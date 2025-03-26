import requests
import lxml.html
import zipfile
import shutil
import os
import glob
import logging
import sys
import time
from datetime import date


def cas_login(service, username, password):
    try:
        # GET parameters - URL we'd like to log into.
        params = {'service': service}
        LOGIN_URL = 'http://192.168.0.22:8080/cas/login'

        # Start session and get login form.
        session = requests.session()
        login = session.get(LOGIN_URL, params=params)

        # Get the hidden elements and put them in our form.
        login_html = lxml.html.fromstring(login.text)
        hidden_elements = login_html.xpath('//form//input[@type="hidden"]')
        form = {x.attrib['name']: x.attrib['value'] for x in hidden_elements}

        # "Fill out" the form.
        form['username'] = username
        form['password'] = password

        # Finally, login and return the session.
        session.post(LOGIN_URL, data=form, params=params)
        #logging.debug("TEST")
        return session
    except Exception as Argument:
        logging.exception(f"{Argument} in cas_login module")


def export(wnids):
    results = []
    try:
        with cas_login("http://192.168.0.22:8082/SmartCMS/websites-overview", "fnc_scr", "gf&UYfh^&48") as s:
            for wnid in wnids:
                a = s.post("http://192.168.0.22:8082/SmartCMS/get-website-from-cms", params={"wnid": wnid})
                results.append(a)
                #print(a.json())
                #print(a.json().get('url'))
            s.close()
        return results
    except Exception as Argument:
        logging.exception(f"{Argument}")


def download(ress):
    try:
        for res in ress:
            url = res.json().get("url")
            name = res.json().get("name")
            response = requests.get(url)
            open(f"{name}.zip", "wb").write(response.content)
    except Exception as Argument:
        logging.exception(f"{Argument} in export module")


def unzip_and_copy(ress):
    # target = "c:/temp/test"
    now = time.strftime("%Y-%m-%d_%H%M%S")
    targets = [f"//192.168.0.20/common/slajdy/aktualne/{now}", f"//192.168.168.4/wspolny/slajdy/aktualne/{now}"]
    #target = "//192.168.0.20/common/slajdy"

    for res in ress:
        try:
            name = res.json().get("name")
            print(name)
            zip_file = f"{name}.zip"
            # folder = 'unknown'
            # if "klippan" in zip_file.lower():
            #     print("klippan")
            #     folder = "klippan"
            # elif "andrenplast" in zip_file.lower():
            #     print("andrenplast")
            #     folder = "andrenplast"
            for target in targets:
                time.sleep(5)
                if not os.path.exists(target):
                    time.sleep(5)
                    os.makedirs(target)
                if not os.path.exists(f"{target}/{name}"):
                    time.sleep(5)
                    os.makedirs(f"{target}/{name}")

                files = glob.glob(f"{target}/{name}/*")
                for f in files:
                    os.remove(f)

                with zipfile.ZipFile(zip_file, mode="r") as archive:
                    for file in archive.namelist():
                        if file.startswith("sites/default/files/tpvision/") and (file != "sites/default/files/tpvision/"):
                            archive.extract(file, target)
                            shutil.move(f"{target}/{file}", f"{target}/{name}")
                    shutil.rmtree(f"{target}/sites")
        except Exception as Argument:
            logging.exception(f"{Argument} in unzip_and_copy module")


def delete_old():
    try:
        folders = ["//192.168.0.20/common/slajdy/aktualne", "//192.168.168.4/wspolny/slajdy/aktualne"]
        for folder in folders:
            subfolders = [f.path for f in os.scandir(folder) if f.is_dir()]

            for subfolder in subfolders:
                shutil.rmtree(subfolder)
    except Exception as Argument:
        logging.exception(f"{Argument} in delete_old")


logfile = "//192.168.0.20/common/slajdy/log.txt"
if os.path.exists(logfile):
    pass
else:
    open(logfile, "w")
logging.basicConfig(filename=logfile,
                    filemode='a',
                    format='%(asctime)s,%(msecs)d %(levelname)s %(message)s',
                    datefmt="%Y-%m-%d %H:%M:%S",
                    level=logging.DEBUG)

logging.info("")
logging.info("----------------------------")
logging.info("********** START **********")

delete_old()
r = export([45, 92, 1815])
download(r)
unzip_and_copy(r)


