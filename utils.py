
import os
from subprocess import run
from constants import WIN_ROOT


def disable_password_expiration(user: str):
    run(f"{os.path.join(WIN_ROOT, 'wbem', 'WMIC.exe')} useraccount where \"Name=\'{user}\'\" set PasswordExpires=FALSE")
