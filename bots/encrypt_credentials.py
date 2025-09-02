#!/usr/bin/env python3
"""
Credential Encryption Utility for NA Contract Numbers Search Bot
This script helps encrypt your ERP credentials for secure storage in .env file
"""

from cryptography.fernet import Fernet
import os
from pathlib import Path

def generate_key():
    """Generate a new encryption key"""
    print("=== Generating New Encryption Key ===")
    key = Fernet.generate_key()
    key_str = key.decode()
    
    print(f"🔑 Generated Key: {key_str}")
    print("\n⚠️  IMPORTANT: Save this key securely!")
    print("   - Store it in a separate secure location")
    print("   - Don't commit it to version control")
    print("   - You'll need it to decrypt credentials")
    
    # Save key to .key file
    root_dir = Path(__file__).parent.parent
    key_file = root_dir / ".key"
    
    with open(key_file, 'w') as f:
        f.write(key_str)
    
    print(f"\n✅ Key saved to: {key_file}")
    print("   (This file should be kept secure and not shared)")
    
    return key_str

def encrypt_credentials():
    """Encrypt username and password"""
    print("\n=== Encrypting Credentials ===")
    
    # Check if .key file exists
    root_dir = Path(__file__).parent.parent
    key_file = root_dir / ".key"
    
    if not key_file.exists():
        print("❌ No encryption key found!")
        print("Please run the key generation first or create a .key file manually.")
        return
    
    # Load the key
    with open(key_file, 'r') as f:
        key_str = f.read().strip()
    
    try:
        key = key_str.encode()
        fernet = Fernet(key)
        
        print("Enter your ERP credentials to encrypt:")
        username = input("Username: ").strip()
        password = input("Password: ").strip()
        
        if not username or not password:
            print("❌ Username and password cannot be empty!")
            return
        
        # Encrypt credentials
        enc_username = fernet.encrypt(username.encode()).decode()
        enc_password = fernet.encrypt(password.encode()).decode()
        
        print("\n=== Encrypted Credentials ===")
        print(f"Encrypted Username: {enc_username}")
        print(f"Encrypted Password: {enc_password}")
        
        # Create .env content
        env_content = f"""# Encrypted ERP System Credentials
# DO NOT commit this file to version control
# Use the .key file to decrypt these values

ENC_USERNAME={enc_username}
ENC_PASSWORD={enc_password}

# Note: Keep both .env and .key files secure
"""
        
        # Write .env file
        env_file = root_dir / ".env"
        with open(env_file, 'w', encoding='utf-8') as f:
            f.write(env_content)
        
        print(f"\n✅ Encrypted credentials saved to: {env_file}")
        print("🔒 Your credentials are now encrypted and secure!")
        
    except Exception as e:
        print(f"❌ Encryption failed: {e}")

def decrypt_test():
    """Test decryption of credentials"""
    print("\n=== Testing Decryption ===")
    
    # Check if both files exist
    root_dir = Path(__file__).parent.parent
    key_file = root_dir / ".key"
    env_file = root_dir / ".env"
    
    if not key_file.exists() or not env_file.exists():
        print("❌ Missing .key or .env file!")
        return
    
    try:
        # Load key
        with open(key_file, 'r') as f:
            key_str = f.read().strip()
        
        # Load encrypted credentials
        with open(env_file, 'r') as f:
            env_content = f.read()
        
        # Parse .env file
        env_vars = {}
        for line in env_content.split('\n'):
            if '=' in line and not line.startswith('#'):
                key, value = line.split('=', 1)
                env_vars[key.strip()] = value.strip()
        
        enc_username = env_vars.get('ENC_USERNAME')
        enc_password = env_vars.get('ENC_PASSWORD')
        
        if not enc_username or not enc_password:
            print("❌ Encrypted credentials not found in .env file!")
            return
        
        # Decrypt
        key = key_str.encode()
        fernet = Fernet(key)
        
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()
        
        print("✅ Decryption successful!")
        print(f"Username: {username}")
        print(f"Password: {'*' * len(password)}")
        
    except Exception as e:
        print(f"❌ Decryption failed: {e}")

def main():
    """Main menu"""
    print("=== NA Contract Numbers Search Bot - Credential Encryption ===")
    print("\nChoose an option:")
    print("1. Generate new encryption key")
    print("2. Encrypt credentials")
    print("3. Test decryption")
    print("4. Exit")
    
    while True:
        choice = input("\nEnter your choice (1-4): ").strip()
        
        if choice == '1':
            generate_key()
        elif choice == '2':
            encrypt_credentials()
        elif choice == '3':
            decrypt_test()
        elif choice == '4':
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please enter 1-4.")

if __name__ == "__main__":
    main()
