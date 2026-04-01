import getpass
from werkzeug.security import generate_password_hash

password = getpass.getpass('Password: ')
print(generate_password_hash(password))
