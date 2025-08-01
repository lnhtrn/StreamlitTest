import secrets
cookie_secret = secrets.token_hex(32)  # Generates a 64-character hexadecimal string
print(cookie_secret)