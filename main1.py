import msoffcrypto
import threading
import argparse
import time
from pathlib import Path
from typing import List, Tuple, Optional
from tqdm import tqdm

def read_wordlist(wordlist_path: str) -> List[str]:
    """Read passwords from wordlist file"""
    with open(wordlist_path, 'r', encoding='utf-8', errors='ignore') as f:
        return [line.strip() for line in f if line.strip()]

def try_password(file_path: str, password: str) -> bool:
    """Try a single password"""
    try:
        with open(file_path, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            # Check if file is encrypted
            if not office_file.is_encrypted():
                raise ValueError("File is not encrypted")
                
            # Try different encodings
            for encoding in ['utf-8', 'ascii', 'latin1']:
                try:
                    # Pass the password directly to load_key
                    office_file.load_key(password=password)
                    
                    # Create a temporary file-like object for decryption
                    from io import BytesIO
                    decrypted = BytesIO()
                    office_file.decrypt(decrypted)
                    return True
                except msoffcrypto.exceptions.InvalidKeyError:
                    continue
                except Exception as e:
                    print(f"Error with encoding {encoding}: {str(e)}")
                    continue
            return False
            
    except Exception as e:
        print(f"Error trying password '{password}': {str(e)}")
        return False

def crack_single_thread(file_path: str, passwords: List[str]) -> Tuple[Optional[str], float]:
    """Single-threaded password cracking"""
    start_time = time.time()
    
    with tqdm(total=len(passwords), desc="Testing passwords") as pbar:
        for password in passwords:
            if try_password(file_path, password):
                return password, time.time() - start_time
            pbar.update(1)
    
    return None, time.time() - start_time

def crack_multi_thread(file_path: str, passwords: List[str], num_threads: int) -> Tuple[Optional[str], float]:
    """Multi-threaded password cracking"""
    start_time = time.time()
    found_password = {'value': None}
    chunk_size = len(passwords) // num_threads
    threads = []
    lock = threading.Lock()

    def worker(password_chunk: List[str], pbar: tqdm):
        for password in password_chunk:
            if found_password['value']:
                break
            if try_password(file_path, password):
                with lock:
                    found_password['value'] = password
                break
            with lock:
                pbar.update(1)

    # Split passwords into chunks
    password_chunks = [passwords[i:i + chunk_size] for i in range(0, len(passwords), chunk_size)]
    
    with tqdm(total=len(passwords), desc="Testing passwords") as pbar:
        # Create and start threads
        for chunk in password_chunks:
            thread = threading.Thread(target=worker, args=(chunk, pbar))
            threads.append(thread)
            thread.start()

        # Wait for all threads to complete
        for thread in threads:
            thread.join()

    return found_password['value'], time.time() - start_time

def crack_word_password(
    file_path: str,
    wordlist_path: str,
    mode: str = "single",
    threads: int = 4
) -> Tuple[Optional[str], float]:
    """
    Main password cracking function
    
    Args:
        file_path: Path to the encrypted Word document
        wordlist_path: Path to the password list file
        mode: 'single' or 'multi' for threading mode
        threads: Number of threads to use in multi-thread mode
    
    Returns:
        Tuple of (found_password, time_taken)
    """
    # Validate inputs
    if not Path(file_path).exists():
        raise FileNotFoundError(f"Word file not found: {file_path}")
    if not Path(wordlist_path).exists():
        raise FileNotFoundError(f"Wordlist not found: {wordlist_path}")

    # Load passwords
    passwords = read_wordlist(wordlist_path)
    print(f"Loaded {len(passwords)} passwords from dictionary")
    print(f"First few passwords: {passwords[:5]}")  # Debug line

    # Choose cracking method
    if mode == "single":
        return crack_single_thread(file_path, passwords)
    else:
        return crack_multi_thread(file_path, passwords, threads)

def main():
    parser = argparse.ArgumentParser(description="Microsoft Word Password Cracker")
    parser.add_argument("file", help="Path to encrypted Word document")
    parser.add_argument("wordlist", help="Path to password list")
    parser.add_argument("-m", "--mode", choices=["single", "multi"], 
                       default="single", help="Cracking mode (default: single)")
    parser.add_argument("-t", "--threads", type=int, default=4,
                       help="Number of threads for multi mode (default: 4)")
    
    args = parser.parse_args()
    
    print(f"Starting password recovery for: {args.file}")
    try:
        password, duration = crack_word_password(
            args.file,
            args.wordlist,
            args.mode,
            args.threads
        )
        
        if password:
            print(f"\nPassword found: {password}")
            print(f"Time taken: {duration:.2f} seconds")
        else:
            print("\nPassword not found")
            
    except Exception as e:
        print(f"Critical error: {str(e)}")

if __name__ == "__main__":
    main()