# split 0
#!/usr/bin/python
# -*- coding: utf-8 -*-
from json import dump as PKL_dump, load as PKL_load
import sqlite3
import inspect
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
import logging
from lxml import etree
import distutils.dir_util
import string
import filecmp
from pathlib2 import Path
import shutil
import datetime
import csv
import dicttoxml
from tempfile import gettempdir
import configparser
from PyPDF2 import PdfFileReader, PdfFileWriter
from fpdf import FPDF
from zlib import decompress
from base64 import b64decode
from io import BytesIO
import subprocess
try:
    from subprocess import CompletedProcess
except ImportError:
    # Python 2

    class CompletedProcess:
        def __init__(self, args, returncode, stdout=None, stderr=None):
            self.args = args
            self.returncode = returncode
            self.stdout = stdout
            self.stderr = stderr

        def check_returncode(self):
            if self.returncode != 0:
                err = subprocess.CalledProcessError(self.returncode, self.args, output=self.stdout)
                raise err
            return self.returncode

    def sp_run(*popenargs, **kwargs):
        input = kwargs.pop("input", None)
        check = kwargs.pop("handle", False)
        if input is not None:
            if 'stdin' in kwargs:
                raise ValueError('stdin and input arguments may not both be used.')
            kwargs['stdin'] = subprocess.PIPE
        process = subprocess.Popen(*popenargs, **kwargs)
        try:
            outs, errs = process.communicate(input)
        except:
            process.kill()
            process.wait()
            raise
        returncode = process.poll()
        if check and returncode:
            raise subprocess.CalledProcessError(returncode, popenargs, output=outs)
        return CompletedProcess(popenargs, returncode, stdout=outs, stderr=errs)

    subprocess.run = sp_run
    # ^ This monkey patch allows it work on Python 2 or 3 the same way

from subprocess import Popen as SP_Popen, PIPE as SP_PIPE, \
    check_output as SP_check_output, call as Fcall
from wx.grid import Grid as WX_Grid, \
    EVT_GRID_SELECT_CELL as WX_EVT_GRID_SELECT_CELL
from re import sub as re_sub
import wx, wx.adv
from wx.lib.mixins.listctrl import CheckListCtrlMixin, ListCtrlAutoWidthMixin
from wx.lib.embeddedimage import PyEmbeddedImage
from docx2txt import process as DOC2_proc
from threading import Thread
from math import ceil
from time import sleep, ctime, time
from comtypes.client import CreateObject
import os
import argparse

# STATIC VARS

SPLASHER = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAIAAABD1OhwAAAAA3NCSVQICAjb4U/gAAAgAElE"
")

OICC_LOGO = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "2pOGneYd8Nq8IIL4e4abaS6aAQD+AcEDXTQ45NoDAAAAAElFTkSuQmCC")

OICC_LOGO_2 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "Y84AAAAASUVORK5CYII=")

OICC_LOGO_3 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "olwAAAAASUVORK5CYII=")

OICC_LOGO_4 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AElFTkSuQmCC")

OICC_LOGO_5 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "SUVORK5CYII=")



VERSIONING = 5.9
FILEVERSIONING = str(str(VERSIONING) + ".0.0")

# END OF STATIC VARS

parser = argparse.ArgumentParser()
parser.add_argument("--version", help="report version number", action="store_true")
parser.add_argument("--firstset", help="set version on network", action="store_true")

options = parser.parse_args()
if options.version:
    print(VERSIONING)
    os.sys.exit()

if options.firstset:
    cbconfig = configparser.ConfigParser()
    cbconfig.read(r"..\settings.ini")
    cbserver_file = r'Codebase\server_settings.ini'
    cbpaff = cbconfig['rootpath']['path']
    cbrootpath = cbpaff + '\\'
    cbserver_side = os.path.abspath(cbrootpath + cbserver_file)
    srvconfig = configparser.ConfigParser()
    srvconfig.read(cbserver_side)
    try:
        srvconfig.add_section('info')
        with open(cbserver_side, 'w') as srvconfigfile:
            srvconfig.write(srvconfigfile)
    except:
        pass
    try:
        srvconfig.set('info', 'sw_version', str(VERSIONING))
        srvconfig.set('info', 'file_version', str(FILEVERSIONING))
        with open(cbserver_side, 'w') as srvconfigfile:
            srvconfig.write(srvconfigfile)
        print(VERSIONING)
        print(FILEVERSIONING)
    except:
        pass
    finally:
        os.sys.exit()

config = configparser.ConfigParser()
config.read('settings.ini')
try:
    localversion = (config['info']['sw_version'])
    localversion = float(localversion)
except KeyError:
    localversion = 0
    try:
        config.add_section('info')
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)
    except:
        pass
if localversion < VERSIONING:
    # write VERSIONING to config file
    try:
        config.set('info', 'sw_version', str(VERSIONING))
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)
    except Exception as e:
        print('oops ' + str((inspect.stack()[0][2])))
        print (e.message, e.args)
        print("config oops")
        sleep(4)
        pass

print('please wait while the program loads (this window will minimise shortly)')
logger = logging.getLogger('OICC')
print('modules loaded...')

PRINTABLE = set(string.printable)


def progressbar(it, prefix="", size=60, out=os.sys.stdout):
    count = len(it)

    def show(j):
        x = int(size*j/count)
        out.write("%s[%s%s] %i/%i\r" % (prefix, u"#"*x, "."*(size-x), j, count))
        out.flush()
    show(0)
    for i, item in enumerate(it):
        yield item
        show(i+1)
    out.write("\n")
    out.flush()


def run_once():
    # Code for something you only want to execute once
    timelog = make_safe_filename(str(ctime(time())))
    log_file = os.path.abspath(
        wordy.rootpath + '\\' + 'finalised' + '\\' + 'finalised_' + timelog + '.log')
    try:
        logger.propagate = False
        if not logger.handlers:
            logger.setLevel(logging.DEBUG)
            formatter = logging.Formatter('%(message)s')
            ch = logging.StreamHandler()
            ch.setFormatter(formatter)
            fh = logging.FileHandler(log_file)
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(formatter)
            logger.addHandler(ch)
            logger.addHandler(fh)
    except Exception as e:
        print('oops ' + str((inspect.stack()[0][2])))
        print (e.message, e.args)
    run_once.func_code = (lambda: None).func_code


def set_me(which_frame):
    config = configparser.ConfigParser()
    config.read('settings.ini')
    dlg = wx.TextEntryDialog(which_frame, 'Enter your name',
                             'Enter the name you wish to associate with "me"')
    dlg.ShowModal()
    if len(dlg.GetValue()) > 0:
        try:
            config.set('user_profile', 'me', safestr(dlg.GetValue()))
            with open('settings.ini', 'w') as configfile:
                config.write(configfile)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            print("config oops")
            pass


def opno2proj(stringin):
    y = stringin.find('-')
    x = stringin.find('-', y + 1)
    proj = stringin[y + 1:x]
    return proj


def make_safe_filename(s):
    def safe_char(c):
        if c.isalnum():
            return c
        else:
            return "_"
    return "".join(safe_char(c) for c in s).rstrip("_")


def returnNotMatches(a, b):
    return [[x for x in a if x not in b], [x for x in b if x not in a]][0]


def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)


def safestr(x, replace=False):
    try:
        try:
            p = str(x)
        except UnicodeEncodeError:
            print('filtering non ASCII characters')
            newx = filter(lambda f: f in PRINTABLE, x)
            p = str(newx)
            # p = ''
            # for z in x:
            #    if isinstance(z, basestring):
            #        z = unicode(z).encode('ascii', 'ignore')
            #        p = p + z
        if replace:
            mystring3 = p + " "
            mystring2 = " " + p
            newstring = ""
            for xx, yy in zip(mystring3, mystring2):
                if xx == "&":
                    if yy != xx:
                        newstring = newstring + xx
                else:
                    newstring = newstring + xx
            return(newstring.strip())
            return str(p)
        else:
            return str(p)
    except Exception as e:
        print('oops ' + str((inspect.stack()[0][2])))
        print (e.message, e.args)


def getRnP(inf, short=None):
    lh = inf.rfind('_')
    lx = foolproof_finder(inf, lh)
    lw = inf.find('-')
    REV = inf[lh + 1:len(inf) - 4]
    if short is not None:
        PEOI = inf[lw + 1:len(inf) - lx]
    else:
        PEOI = inf[:len(inf) - lx]
    return (REV, PEOI)


def foolproof_finder(inf, lh):
    lx = len(inf) - lh
    return lx


def search_file(directory=None, file=None):
    assert os.path.isdir(directory)
    (current_path, directories, files) = os.walk(directory).next()
    if file in files:
        if safestr(directory).find(wordy.forappralpath) != -1:
            return ('AW. Approval', os.path.join(directory, file))
        elif safestr(directory).find(wordy.approvedpath) != -1:
            return ('Approved', os.path.join(directory, file))
        elif safestr(directory).find(wordy.archive_path) != -1:
            return ('Archived', os.path.join(directory, file))
        else:
            return ('Error', os.path.join(directory, file))
    elif directories == '':
        return None
    else:
        try:
            for new_directory in directories:
                result = search_file(directory=os.path.join(directory,
                                                            new_directory), file=file)
                if result:
                    return result
            return None
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            return ('Not Found', None)


def findingNemo(xl1, REV):
    for p in range(1, len(xl1)):
        if safestr(xl1[p]['revs']['rev']) == REV:
            found2 = p
            skippy = False
            break
        else:
            found2 = len(xl1)
            skippy = True
    return [found2, skippy]


def findme(xl2, PEOI):
    for x in range(0, len(xl2)):
        try:
            if safestr(xl2[x][0]['details']['opno']) == PEOI:
                return x
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            for opin in xl2[x][0]['details'].items():
                if safestr(opin[1]) == PEOI:
                    return x
    return -1


def loadUP(dBf):
    if not os.path.isfile(dBf):
        dlg = wx.MessageDialog(None, 'Cannot find database file!', '',
                               wx.OK | wx.ICON_ERROR)
        dlg.ShowModal()
        dlg.Destroy()
        return (False, None)
    with open(dBf, 'rb') as f:
        x = PKL_load(f)
    return (True, x)


def moreFinding(fd, rv):
    if len(fd) > 0:
        findme2 = findingNemo(fd, rv)
        return findme2[0]
    else:
        return -1


def gen_obs():
    OBS_string = r"JVBERi0xLjUKJbXtrvsKNCAwIG9iago8PCAvTGVuZ3RoIDUgMCBSCiAgIC9GaWx0ZXIgL0ZsYXRlRGVjb2RlCj4+CnN0cmVhbQp4nDNUMABCXUMgYWppqmdgYWBgaK6QnMtVCISGYEkICRTSTzRQSC9W0K8wU3DJ5wrEo8CckAILQgosCSkwNCCowpCgCiOCKowh"
    return (b64decode(OBS_string))


def fileaccesscheck(file2open):
    try:
        if os.path.isfile(file2open):
            f = open(file2open, 'rb')
            f.close()
            tempname = os.path.abspath(
                wordy.rootpath + '//' + safestr(map(ord, os.urandom(1))[0]) + ".tst")
            shutil.move(file2open, tempname)
            sleep(0.2)
            shutil.move(tempname, file2open)
            sleep(0.2)
            return True
        else:
            return False
    except IOError:
        return False
    return
    
