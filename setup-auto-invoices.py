def sha256sum(filename):
    h  = hashlib.sha256()
    mv = memoryview(bytearray(128*1024))
    with open(filename, 'rb', buffering=0) as f:
        for n in iter(lambda : f.readinto(mv), 0):
            h.update(mv[:n])
    return h.hexdigest()

def update(req, localFilePath):
    with open(localFilePath,'wb') as f:
        f.write(req.content)

def checkForUpdates(localFilePath, downloadLink):
    try:
        req = requests.get(downloadLink)
    except Exception:
        print("No internet connexion. Cannot install program...")
        return
    if os.access(localFilePath, os.F_OK) is not True:
        print("Installing program...")
        update(req, localFilePath)
        print("Installation complete!")
        return
    if "\n" not in req.text:
        print("Corrupt online file!")
        return
    fl_online = req.text[0:req.text.index("\n")]
    with open(localFilePath) as f:
        fl_local = f.readline().strip()
    if fl_local == fl_online:
        if hashlib.sha256(req.content).hexdigest() == sha256sum(localFilePath):
            print("Latest file! No need to update.")
        else:
            print("Corrupt file! Repairing file...")
            update(req, localFilePath)
            print("Update complete!")
            return
    else:
        print("You are not running the latest version. Need to update!")
        update(req, localFilePath)
        print("Update complete!")
        return

def checkRequirements(localFilePath, downloadLink):
    try:
        req = requests.get(downloadLink)
    except Exception:
        print("No internet connexion. Cannot install requirements...")
        return
    if os.access(localFilePath, os.F_OK) is not True:
        print("Installing requirements...")
        update(req, localFilePath)
        print("Installation complete!")
        return
    if hashlib.sha256(req.content).hexdigest() == sha256sum(localFilePath):
        print("Requirements are correctly installed!")
    else:
        print("You do not have the correct requirements installed. Installing...")
        update(req, localFilePath)
        print("Installation complete!")
        return


if __name__ == '__main__':

    import os, sys, hashlib, subprocess
    # install requests
    print("\nSetting up connexion...")
    process = subprocess.run([sys.executable, "-m", "pip", "install", "requests>=2.23.0"], stdout=subprocess.PIPE)
    if process.returncode == 0:
        import requests
        # check auto-invoices.py
        checkForUpdates("auto-invoices.py", "https://raw.githubusercontent.com/1Intuition/Auto-Invoices/master/auto-invoices.py")
        # check requirements.txt
        checkRequirements("requirements.txt", "https://raw.githubusercontent.com/1Intuition/Auto-Invoices/master/requirements.txt")
        # install requirements
        if os.path.isfile("requirements.txt"):
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], stdout=subprocess.PIPE)
        else:
            print("Could not check requirements! Program might not work correctly...")
    else:
        print("No internet connexion! Cannot update or install program and requirements.")

    exec(open("auto-invoices.py").read())