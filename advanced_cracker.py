import msoffcrypto
import threading
import argparse
import time
import logging
import sys
import os
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
from tqdm import tqdm
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor
from enum import Enum
import tempfile
import hashlib

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('password_cracker.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

class CrackingMode(Enum):
    SINGLE = "single"
    MULTI = "multi"
    HYBRID = "hybrid"

@dataclass
class CrackingStats:
    attempts: int = 0
    start_time: float = 0.0
    end_time: float = 0.0
    found_password: str = ""
    success: bool = False
    errors: List[str] = None

    def __post_init__(self):
        self.errors = []

class PasswordCracker:
    def __init__(self, 
                 file_path: str, 
                 wordlist_path: str,
                 mode: CrackingMode = CrackingMode.SINGLE,
                 threads: int = 4,
                 chunk_size: int = 1000,
                 timeout: int = 3600,
                 verify_hash: bool = True):
        self.file_path = Path(file_path)
        self.wordlist_path = Path(wordlist_path)
        self.mode = mode
        self.threads = max(1, min(threads, 32))  # Limit threads between 1 and 32
        self.chunk_size = chunk_size
        self.timeout = timeout
        self.verify_hash = verify_hash
        self.stats = CrackingStats()
        self._validate_inputs()

    def _validate_inputs(self) -> None:
        """Validate input files and parameters"""
        if not self.file_path.exists():
            raise FileNotFoundError(f"Document not found: {self.file_path}")
        if not self.wordlist_path.exists():
            raise FileNotFoundError(f"Wordlist not found: {self.wordlist_path}")
        if self.file_path.stat().st_size == 0:
            raise ValueError("Document is empty")
        if self.wordlist_path.stat().st_size == 0:
            raise ValueError("Wordlist is empty")

    def _calculate_file_hash(self) -> str:
        """Calculate document hash for verification"""
        sha256_hash = hashlib.sha256()
        with open(self.file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()

    def _try_decrypt(self, password: str, office_file: msoffcrypto.OfficeFile) -> bool:
        """Attempt to decrypt with a single password"""
        try:
            for encoding in ['utf-8', 'latin1', 'ascii', 'cp1252']:
                try:
                    office_file.load_key(password=password)
                    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                        office_file.decrypt(temp_file)
                        if self.verify_hash:
                            # Verify decryption was successful
                            if os.path.getsize(temp_file.name) > 0:
                                return True
                except msoffcrypto.exceptions.InvalidKeyError:
                    continue
                except Exception as e:
                    self.stats.errors.append(f"Error with {encoding}: {str(e)}")
            return False
        finally:
            if 'temp_file' in locals():
                os.unlink(temp_file.name)

    def _worker(self, passwords: List[str], progress: tqdm, 
                shared_data: Dict[str, Any], lock: threading.Lock) -> None:
        """Worker thread for password testing"""
        try:
            with open(self.file_path, 'rb') as f:
                office_file = msoffcrypto.OfficeFile(f)
                for password in passwords:
                    if shared_data['found']:
                        break
                    
                    if self._try_decrypt(password, office_file):
                        with lock:
                            shared_data['found'] = True
                            shared_data['password'] = password
                            progress.write(f"\nPassword found: {password}")
                        break
                    
                    with lock:
                        self.stats.attempts += 1
                        progress.update(1)
                        
        except Exception as e:
            logging.error(f"Worker thread error: {str(e)}")
            self.stats.errors.append(str(e))

    def crack(self) -> Tuple[Optional[str], float, CrackingStats]:
        """
        Main password cracking method
        
        Returns:
            Tuple[Optional[str], float, CrackingStats]: 
                (found_password, time_taken, statistics)
        """
        self.stats = CrackingStats()
        self.stats.start_time = time.time()
        
        try:
            # Read and preprocess passwords
            passwords = self._load_passwords()
            if not passwords:
                raise ValueError("No valid passwords in wordlist")

            logging.info(f"Starting {self.mode.value} mode with {len(passwords)} passwords")
            
            if self.mode == CrackingMode.SINGLE:
                result = self._crack_single(passwords)
            elif self.mode == CrackingMode.MULTI:
                result = self._crack_multi(passwords)
            else:
                result = self._crack_hybrid(passwords)

            self.stats.end_time = time.time()
            self.stats.success = bool(result)
            self.stats.found_password = result or ""

            return result, self.stats.end_time - self.stats.start_time, self.stats

        except Exception as e:
            logging.error(f"Cracking failed: {str(e)}")
            self.stats.errors.append(str(e))
            return None, time.time() - self.stats.start_time, self.stats

    def _load_passwords(self) -> List[str]:
        """Load and preprocess passwords from wordlist"""
        passwords = []
        with open(self.wordlist_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                password = line.strip()
                if password and len(password) >= 1:
                    passwords.append(password)
        return passwords

    def _crack_single(self, passwords: List[str]) -> Optional[str]:
        """Single-threaded cracking implementation"""
        shared_data = {'found': False, 'password': None}
        with tqdm(total=len(passwords), desc="Testing passwords") as progress:
            self._worker(passwords, progress, shared_data, threading.Lock())
        return shared_data['password']

    def _crack_multi(self, passwords: List[str]) -> Optional[str]:
        """Multi-threaded cracking implementation"""
        shared_data = {'found': False, 'password': None}
        lock = threading.Lock()
        
        # Split passwords into chunks
        chunks = [passwords[i:i + self.chunk_size] 
                 for i in range(0, len(passwords), self.chunk_size)]
        
        with tqdm(total=len(passwords), desc="Testing passwords") as progress:
            with ThreadPoolExecutor(max_workers=self.threads) as executor:
                futures = [
                    executor.submit(self._worker, chunk, progress, shared_data, lock)
                    for chunk in chunks
                ]
                
                # Wait for either completion or timeout
                for future in futures:
                    try:
                        future.result(timeout=self.timeout/len(futures))
                    except Exception as e:
                        logging.error(f"Thread error: {str(e)}")
                        self.stats.errors.append(str(e))
                        
        return shared_data['password']

    def _crack_hybrid(self, passwords: List[str]) -> Optional[str]:
        """Hybrid approach using both methods"""
        # Try common passwords single-threaded first
        common_passwords = passwords[:100]
        result = self._crack_single(common_passwords)
        if result:
            return result
            
        # Then try remaining passwords multi-threaded
        remaining_passwords = passwords[100:]
        return self._crack_multi(remaining_passwords)

# Example usage:
def main():
    parser = argparse.ArgumentParser(description="Advanced Word Password Cracker")
    parser.add_argument("file", help="Path to encrypted Word document")
    parser.add_argument("wordlist", help="Path to password list")
    parser.add_argument("-m", "--mode", 
                       choices=[m.value for m in CrackingMode],
                       default="single", 
                       help="Cracking mode")
    parser.add_argument("-t", "--threads", type=int, default=4,
                       help="Number of threads (1-32)")
    parser.add_argument("--chunk-size", type=int, default=1000,
                       help="Password chunk size for threading")
    parser.add_argument("--timeout", type=int, default=3600,
                       help="Timeout in seconds")
    parser.add_argument("--no-verify", action="store_false",
                       dest="verify",
                       help="Skip successful decryption verification")
    
    args = parser.parse_args()
    
    try:
        cracker = PasswordCracker(
            args.file,
            args.wordlist,
            mode=CrackingMode(args.mode),
            threads=args.threads,
            chunk_size=args.chunk_size,
            timeout=args.timeout,
            verify_hash=args.verify
        )
        
        password, duration, stats = cracker.crack()
        
        if password:
            print(f"\nSuccess! Password found: {password}")
            print(f"Time taken: {duration:.2f} seconds")
            print(f"Attempts: {stats.attempts}")
        else:
            print("\nPassword not found")
            if stats.errors:
                print("Errors encountered:")
                for error in stats.errors:
                    print(f"- {error}")
                    
    except Exception as e:
        print(f"Critical error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()