def FTP_Auth():
    Host = None
    Login = None
    Password = None

    with open("SP_FTP_Host.txt") as I:
        Host = I.read().strip()
    with open("SP_FTP_Login.txt") as I:
        Login = I.read().strip()
    with open("SP_FTP_Password.txt") as I:
        Password = I.read().strip()

    return Host, Login, Password