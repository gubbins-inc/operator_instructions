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
    "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"
    "u//++zmOW7t27a9//euamhoASCaT995779GjRwGgpKTkF7/4BTtlZGTkvvvum56eHhwcLC0t"
    "tdvt3/3ud5cvXw4Ae/bsueuuu2w2G6V07dq1//qv/2p+aZs3b37ooYe2bNmSzWZ37Nhx9913"
    "s54HgM7OzrvuugsABEH4/ve/zy4Vj8cfeOCB48ePj4+Pu1wur9f72c9+9j3veQ875Wc/+9kj"
    "jzySSCQqKip++ctfmp8Ba3NnZycA3HjjjXfffTfb/tprr91yyy033XTTF77whb6+vptuuunw"
    "4cPz+mfxvmP2Zl3XFUX5zne+w6Say+VSFGXdunUVFRXM6GYYxvj4eNfRo5lsVlVVQggAcjod"
    "fr9vzZo1K1eutFqtLERp4S1M70w4HD7WfWxiYkJV1IA10ORrusB1QaW18s+xPx+FoxEpEsvF"
    "hIggIpFEiGvSVUJKvH6v4BWQOEvoaDTKfuzatevVV199+umns9nsLbfcsnnz5vr6egAwDCMU"
    "Cj311FPBYHDnzp0PP/zw5s2bKyoqAEDTtHQ6/cQTT7BwqNLSUrOFoVDommuuede73vWLX/zi"
    "q1/9aj6hp6en9+3bp+v6wMCASWin0/nEE08MDw/fcsstX/va12pra00JGo/Hv/GNb3z/+9+v"
    "q6u79tpr//jHP1533XVsVyKRaGlpueeeezo6Or761a+2t7eXlZUt+lIYvv71r3/0ox+trq7+"
    "9Kc/bYo6XddjsdiPfvSjm266qa6uLhwOswoQAHDixInnnntu7969qqp+5jOfmZqaYkSvrKz8"
    "5je/qapqXV3dY4891tbWVl1dzU7RNI0Q8qUvfclms33qU5968cUXP/KRj5hdLcsy69JYLGYO"
    "RwDwxBNPtLS0fOlLX/rGN74Rj8fZRq/X++CDD2az2WuuuWbHjh033HBDfg9fffXV8Xj88ccf"
    "f+yxx/JNBbt373711Vf/53/+p6ur63Of+9yNN97IelJV1Wg0+thjj91xxx26rofD4YX9c3Zf"
    "HSEknU6zUI3Kyko2ZCCEksnkgQMHfve730m5HMt1ZRK9tramvb2d+QItFsuiugETz9lstrOz"
    "c9euXaFwSFd1h+AodBSu96xvc7XVueo2+zZvLtxc463xWr0AoBBlXB3/fe73+/bvG9o7JA1I"
    "RszIj30lhPzyl7+84oorVq1atWXLFrvd3tHRwXaxDJeampr6+vqNGzcmEon8NwEALCeXDUr5"
    "24uLiy+88ML169ebQ7PZ44yar7zySv72ioqK8vJyQRBqa2srKytNwr388sslJSWrV69ubGzc"
    "unXrb37zm/yzvF7v2rVrL730UpaftvTrKC0tZVqfGSpjoqmp6aGHHpq3sbCw0OFwvPjii36/"
    "/8UXX2RsZn1SWlpaXV0tCEJVVZXJZgZBEKqrq1tbWysqKjKZzNJNYnjllVduuummkpKSb37z"
    "mxs3bjTvUlBQUF1dbbVaS0pKqqur83vY5XIFAgFRFJlwMfHb3/720ksvbWpquuSSSyoqKn7/"
    "+9+buwoKCgKBwDe/+U0WQbSwGWc32zFdqry8fP369Rs2bAgGg6IoKooyNTX16quvHjlyhOWt"
    "chxntdrWrVuzcuXKsrKyM1EZ8kza8Xj88OHDg4ODqWRKwILP5qtyVi13LC91lfKYbwu0lVpL"
    "rci6l+5N4qRqqDmSIyrZa+zlujmn3enDPpvHBnPhpblcLhwOX3nllSzitKWlZWxszHwEjPEN"
    "N9zgdrsnJyd37NiRLwUlSfqnf/onloXwwgsv5Df1O9/5zjPPPBONRj/5yU/mt3/Xrl1XXXVV"
    "cXHxb37zG03T8jMPFk0PGx0draqqYiPVihUrnnnmmfy9L7744vbt28fGxnbs2JGv379V3Hbb"
    "bV/+8pf37t2bv7GiouKxxx77yle+8qtf/eq+++674oor5p3FHAj5WzDGoVDo9ttvT6VSNptt"
    "06ZNb+buiqIs/MZMcBz35iOBx8fH2QjmdrsrKirGx8fNXS6X67rrrnv00UeXLVtmjkL5OAuh"
    "EUJ2uz0QCLS1tbW0tBQXF7Mk1u7u7oMHD/b390ciEa/XizF2OBx1dTXLly8vLi622Wxn8puw"
    "pENVVcfHx48fPz42OpZMJA3d8Fq8te7a1e7VfpvfwlkQQgInFDoKV2urCZBduV0IkGIoKlER"
    "Rbvl3dpBbbl9eUN9g9lNoig6HA7z3UxOTl522WXmU3Acd+211xYWFhYXF1988cX5csJutz/6"
    "6KMcx80LPOQ4bs2aNePj401NTR//+MfN7eFwuKurq6+vz2q1RqPRoaEhFjOYf695T+3z+dgE"
    "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"
    "68A2wOkAAAAASUVORK5CYII=")

OICC_LOGO = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAALpSURBVDhPVVPdSxRRFD/ztTPbfmKu"
    "piJFbvgRVguLFNFD9S6K2Zdk+C6+VPQQrbIZZPjQfyC9qK0K0osmYulGJeZDBqIF+cHuyM7s"
    "prizO87s7N7mXmfFfnDn/s45v99l5sw9FBzDfN98hx/8gxWFilNiUoSSTyVgy9kgcTkB3jNe"
    "kBl5Ydm13NP2qO2HZYGjA/SQjjiKI1zKSlAWKSO8CLFJhMqGSsJ1pAMf5omXJgnTvMAutM4x"
    "c3c2YGNCMiQ5eTIJikcB1anCgeMAtH0NJJByMSr2OWqLNmshDWEvTA5MBlEvQlt9W19JwgLO"
    "HV9jL8cuWCWI9kUf4txs/2wrHcgG3uCkkBeCpGrhF/1r3KIgI1luf9a+YoVQDdV38V5n1A3S"
    "vrzvqrgiQuZbhp1+Nd1EFCZqe2vb0yidz6EclIXL/muI8le5Ih1IUF4oPwux+zGETiOkNqjk"
    "VeXncsrSQSQSYcxls0JY6117jzW7l3aJJ94dR5RxzkCMzpBm2W/ZiTCP8jDsH3Z29HYoCBCw"
    "GyythJScg3IwuK6Om1rFDml3GmiN04jJYAyyY2QKGb2zszODzYhCMDQ0xPPAEzMGa7Bk5/M8"
    "0Hs39khQPGiVXn3r6ffwmE90T5zg/nBUV1fXARfmqCRKks+jCofXRw7KQOVDeZTSUpBACanx"
    "dWM5qVgwQgbKoqzhfuE+vGEm8G8PfAwsMV4GXH4X0AkqEfcJPhDsQtTSEKRCqX2GMkW0i13t"
    "W31npaHlact3Z5MTqmqrYJfZnaLXbeuPcWGT3hwhChNTA1PXSqgSlxVCPaq/bVEClVLJLCy6"
    "F3vIx2RD2cISt/Sgxqh5Ykf2i8qiApXblbiFYNAGaKwG8ZvxPafb+WWb2R4183owFxxxhB30"
    "0TCpIbUgUAKJxZ/m4CwfDk4RyXtJKOVLCTf7grAZc/LAsIft9Kh3tGELtj64zrsgfj0OuqBD"
    "2pOGneYd8Nq8IIL4e4abaS6aAQD+AcEDXTQ45NoDAAAAAElFTkSuQmCC")

OICC_LOGO_2 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAMDSURBVDhPVZNZSFRRGMf/59w7m86m"
    "ppNJ2GJUZkQgUlS00fIQtm9Ehr0EEb1U9BCNMvVQ0UOPPWUbZWERvWRBZRqRmEEW0oLblBOz"
    "OI7jzB1n5s493XO8E/W7cL/lfP/DPd89H8E/3Gl+c6gEVVftWvnMYCSASEcxSNYM54ogPHPc"
    "UKRwZ9jRe/Loqd2fDAn+btDhzTCJmIQfU0L4+bBM+HnsdQHMrZ4l/BzLYJ3PIrSUv7h4RO7c"
    "OSi92j+OoUdxNRRmJREQVwLEngIpnIISTyOJUDZOfr31m7vqO7xpxrXkxuUntQuU7T0TxP9+"
    "W3PlSp7kdDUxUZCn39y27Ni5vX3cv9vcdaSSrb45ZHq5i5Yqy6/xpJSz1nKbZ4x+bzNcJFk4"
    "nBdzXJh9gNsSddFVWpArXTXQF8Dw+6R8+1J7najQ2dG0cG+aTeZyLIutvrL/GjIRTaycmArB"
    "rnnm0d+DUSgf9eaMlmNuakt3+/nwmFGHaE27JVTzxGKEeNz09Sk/WryzwuVvLUMkGgK5vkBl"
    "WkYSzarZYxOFGsthpOqePdF0KAEwHB+S6UtvImsmhRJf/9KWAkvotc5JUM2UFiJI6rTVSWvJ"
    "TENDQ5KLQRhaWlosMixCzCGqLCzNWUA9G2IiyG8Upv23Nl90ic82n3hUcHzQRBobG6fW+kwk"
    "ySLieEybvj6O2jBIpzfH4ukxJFgwdODKUo9YMXjjVVmGKeqmC87pG6bDf7v6enmPzS3BU+UA"
    "TZDgqMtaCqvN2mXUCJ57x+KUSLBSh/y4uf+BkcbRszs+VNbZMWdhBdLS+DMaNX87zRdidPi+"
    "qNC5ffnZmgJS7DBClLLF+wxXoJKUmIWgs/ukOMwrr6IFTD2Hi9T5Z2RmWzbcrTffP0tvof5Q"
    "FZqcRtHG0ZjDaX8Xl/ytejZTnq29v9FXSP8O02tvSpOJVcQDnwNQeqcHJ0/lwQiclhnC1/vC"
    "uJj74sVZ77PRb+7W6hhGnlcsccC9fhTEmtEHahIz63/DbnYjjsCPAdOL+rwYAP4ARjo/HNF8"
    "Y84AAAAASUVORK5CYII=")

OICC_LOGO_3 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAMDSURBVDhPVZPbTxNBFIfPzO52t9CW"
    "glKgeCGIQbwEIUjES+It8UXxipcQMLwbX9T4YCyk+iCGB/8DNCZSFRJiTECNNzAERIiAENAE"
    "hECxN4RSet3dcWfcKn7J7jln5vwmM2fOIFhFw7MP1ZBd0KSm52T7593wPpIBCWyA3ZwH8mxW"
    "4JZ9XeaZgStXa88M6RL4u4BjKE4QLzA/HPDC02Ub85OUgxu25tmZTxJxcO4UmRbTHxXz412n"
    "uNG358Ez1SYveX1ron5Ik0NgUiOQqkYhGI4BLHoTKDD70TDRXen4EiNUixoftZeFS0/0I99M"
    "b8OBjRV0kFI/SlhCEsOn1uKbdVXD1G9o7b5EivY9EL6+OY3DG0ru00FFkMqoTYLnv7XqLpAl"
    "ny8pZqxdf4Ea2b6lCSumzL3DU27oda/wd12d5SxBo/5IYRUJLytEToBzj+2/giwEQxXRBS+o"
    "lqx8POlZgEFkhzkpByI7jvbd6vEF9DzYPtUpbp9oF/UQ6l+PP6dH61Jy01xBGyz4vIA2v5BJ"
    "HHGsWGfzjSyRKAoUfHlsqs+oDoFWiqnjPHYMhBJISuXofOtkBELYCJbEMmBB1aqrwSkysxQ1"
    "uhKvra1doWKkfc3NzSLwIhNTZOCZVXgR8KGURRYkF8KzYw/vlKexbV8OtKVMVgqorq4u6iwW"
    "EAn62fFU9Kd9yngfIMewQmLBAJBFj/fesR1ZbEbHMSwTEgvLt3dZ/nSYBr32d6kl/VZtQwWZ"
    "ZsAo6JmT0jPBKEndeg7D0RsIIo4DnGLmG16OPdGH4UbNyc/ldhMU5uUCt/KrAxt+TlyjE9j/"
    "o4VlaDS2dOxH5gyzHgJZV3ROdxkoHmFvwTLdd4UdxjEYVoXJ/hrZtuk6MRiL+9whmBFp3xPA"
    "qlYyOQaHhblFi8nUwwVmXFpx44n8shZnaSr+95gGIyoSJRaPaI01oPXGai5a/SBa1zJfqwuh"
    "YuqzH8VZasTWT66t4J1+uS3LDAe5OZDUOKRpd10pzYPBZNVa0P1dGHlVmRQDAPwGnfJClqGB"
    "olwAAAAASUVORK5CYII=")

OICC_LOGO_4 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAL+SURBVDhPTZPZS1RRHMe/59w7m84i"
    "5oylLVKGZkYEIklFVES9ZKstRMYU9BS9VPQQjTL1kOFD/8EULVoYVC8ttGqLYgYVSJta4ozM"
    "4jLjLM7MnXu653hH+jzc3/o9nPO75xD8x9s7rUcrnWhfZFUXBoIRFE+9gZFkEbSuR5GrAuGE"
    "1D0QtJ3Z7z77RZdgfoHMOw8zSHNhaDoJV+C+8PMECupRVlEj/EyOwbTRK5qpSGji7hF576vf"
    "0qGRSTwIRZVwRF2AOHEgBStmSSHSyRhCCWTHouRdzx9jY1rTcC156Gur212d7B+Nkt5lO1sb"
    "eJLDPraIhjxd34xrm05d/Mr9no7W4xsr2I2XQ4Z9dF1p8jpPmmmujts8PyO0S3cRjrNwXsxZ"
    "4sBhbqtLlHbqLMxtCAx/RWK0V35662q96NCo2tXSNJNmuax2Xtd2r0tPC+KxyYZQdBalNnU5"
    "nRwfRtnsZyyifuyoTPWFX1ya0PvwNFRrehioNekhvj9qecyPVp7odrj8nQhFJkEU30omqRkx"
    "LEvNAdGYUxnuDlVaj2Zb4nwQ8skRGn/ryRYaicTrqcEuWBDHDOyg6ZyB56DM1QSJtJppbm5O"
    "cDEjBD6fz2SSMd8gU0VYE82BTju3iiC/0GCI3nRsvSK2/QCnCwwnhonb7Z41bPCSSJyJ4xGm"
    "coNwQR1I7qOHTUTTCM6w0Jqma6WioqO897Bkhin2LZfnVtfgv30dXvdLliLYXJWgwRjxOx1m"
    "mM2WHr1HMPHSE5Mogc1M5cFHrff0NPa4L3yyLq1H+bIqTKWlJ/THhPEcL/yZoh2iQ+PJ7bZN"
    "xQXEpodY5WIHdVeQUoh4C30B+xlxn5PdHrV/zHBsxQLlvEVma+N/+1BGR8G0KSoqRZrJ8Nu2"
    "TVtt9g+jUalTy2fqyrMdhZu92h51Uj0e1SxrI9cIDH/T7saAyOeJLD6CEvvcldDmwriY++LD"
    "sWzy0s7Bopq/03hmK1sNv30LMsSMGe1BjTsbUWQ1IhDDr+c/DY15MQD8AxB+PQYdaDj7AAAA"
    "AElFTkSuQmCC")

OICC_LOGO_5 = PyEmbeddedImage(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1B"
    "AACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAMzSURBVDhPHVNdSJNhFH6/b/861zYH"
    "qcyxQUEq4Y3omDkTTCoUCWSCCYI/IN6Mghp4oYREYHYRQy8MybBQcgq7ELwIE8RwGgjepISz"
    "YLIJY26u6X7c3p7zfXDY3vc75znPeZ7zCfX19RZRFM2IqtbW1p6hoaGmsrIyBc4snU4zhULB"
    "BEFgnHO2v7+fnJmZWT86OlrH+bBQKIQEu93uwMG6srLyRa1WM6VSyWQyGbu8vJT+E1Aul5Oi"
    "qKiIZbNZ6V1XV1cPgP/KKisr23w+36fp6emve3t7v1GkuLq60gJURsXX19dScSwWI+DC8fFx"
    "eGFhYW1qamocTX/Iq6urH2s0GtbZ2en0eDxvtre3/YyxitXV1dfEiOhnMhk2Pj7+MRgM/gSY"
    "vA0P1TQ1NbnE4eHhNpVKxSwWy025XJ5H10g+n/8FoEPSgCIcDl9h7m8ACyLnX3Nz810ab3Bw"
    "8L5YW1urTaVS7Pz8XKipqbFjxlJ0zE5MTLxPJpOc5u3r63uJ+wS6a8CuXK/XV5BOVqtVJeCS"
    "I1HqpNVq2dnZWbq3t/cFWGUQJUgUAJIAgHxsbOy5w+G4jWZSrtFoZCQSBwOOQo7OPJFI8Egk"
    "wjs6OjyBQIBTtLS0PAuFQoV4PM7RiGMkKe/i4oKLUFxSGp5KXlMgKQ/QNOhKD8TE6HJIIEii"
    "0m6QQ5LIEIyTRUBjpaWlbGNj42BycnIaIhVA0wiwDBqkwE4N+95iydSkGQlfUlLCJA1oOWBR"
    "xu12vyIHkJwDE5vf7/fiXcHlco2ASAigegj9EHm9BoOBQUwmnp6eZlHEMHcISRFQTYOiYWlp"
    "6R15rdPpxNHRUTc6KnEfg50BEo+2cnd3Ny7Ozc1tEYPNzc0DUNWh8626urp2k8mkpBkpnE5n"
    "Fdyowrty5MghaJRqvF7vmtDY2DiwuLj4YX5+3t/e3n7PbDabSBOiBzbSGpPQsDIXjUbDW1tb"
    "B2Cc7e/vf9Ld3T0gwFcHHLAuLy9/Bi2BxKG9oF9Sm9yhVS4uLmY0KnUGGEfxU4z8R2hoaLDA"
    "GjPijs1mezAyMvIIn/gN2jSylwpoXio+OTlJz87Oft/Z2fFhHHzOhdB/8jTtaiDwUBoAAAAA"
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
    OBS_string = r"JVBERi0xLjUKJbXtrvsKNCAwIG9iago8PCAvTGVuZ3RoIDUgMCBSCiAgIC9GaWx0ZXIgL0ZsYXRlRGVjb2RlCj4+CnN0cmVhbQp4nDNUMABCXUMgYWppqmdgYWBgaK6QnMtVCISGYEkICRTSTzRQSC9W0K8wU3DJ5wrEo8CckAILQgosCSkwNCCowpCgCiOCKowhKgK5AH8IP3kKZW5kc3RyZWFtCmVuZG9iago1IDAgb2JqCiAgIDgxCmVuZG9iagozIDAgb2JqCjw8CiAgIC9FeHRHU3RhdGUgPDwKICAgICAgL2EwIDw8IC9DQSAwLjUwNyAvY2EgMC41MDcgPj4KICAgPj4KICAgL1hPYmplY3QgPDwgL3g2IDYgMCBSIC94NyA3IDAgUiAveDggOCAwIFIgL3g5IDkgMCBSIC94MTAgMTAgMCBSIC94MTEgMTEgMCBSIC94MTIgMTIgMCBSIC94MTMgMTMgMCBSID4+Cj4+CmVuZG9iagoyIDAgb2JqCjw8IC9UeXBlIC9QYWdlICUgMQogICAvUGFyZW50IDEgMCBSCiAgIC9NZWRpYUJveCBbIDAgMCA4NDEuNjc5OTkzIDU5NS4wODAwMTcgXQogICAvQ29udGVudHMgNCAwIFIKICAgL0dyb3VwIDw8CiAgICAgIC9UeXBlIC9Hcm91cAogICAgICAvUyAvVHJhbnNwYXJlbmN5CiAgICAgIC9JIHRydWUKICAgICAgL0NTIC9EZXZpY2VSR0IKICAgPj4KICAgL1Jlc291cmNlcyAzIDAgUgo+PgplbmRvYmoKNiAwIG9iago8PCAvTGVuZ3RoIDE1IDAgUgogICAvRmlsdGVyIC9GbGF0ZURlY29kZQogICAvVHlwZSAvWE9iamVjdAogICAvU3VidHlwZSAvRm9ybQogICAvQkJveCBbIDAgMCA4NDIgNTk2IF0KICAgL0dyb3VwIDw8CiAgICAgIC9UeXBlIC9Hcm91cAogICAgICAvUyAvVHJhbnNwYXJlbmN5CiAgICAgIC9JIHRydWUKICAgICAgL0NTIC9EZXZpY2VSR0IKICAgPj4KICAgL1Jlc291cmNlcyAxNCAwIFIKPj4Kc3RyZWFtCnic7ZTPipUxDMX3fYo+QWzapE2fQBBcjC7FhYw4IsxidOHr+0u/izDgE8hluPORtvlz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CjAwMDAwMDY0NTIgMDAwMDAgbiAKMDAwMDAwNzAyNCAwMDAwMCBuIAowMDAwMDAxNTg5IDAwMDAwIG4gCjAwMDAwMDE1NjYgMDAwMDAgbiAKMDAwMDAwMjg2NyAwMDAwMCBuIAowMDAwMDAyODQ0IDAwMDAwIG4gCjAwMDAwMDQxMDcgMDAwMDAgbiAKMDAwMDAwNDA4NCAwMDAwMCBuIAowMDAwMDA1MTkwIDAwMDAwIG4gCjAwMDAwMDUxNjcgMDAwMDAgbiAKMDAwMDAwNTczOCAwMDAwMCBuIAowMDAwMDA1NzE1IDAwMDAwIG4gCjAwMDAwMDYzMzcgMDAwMDAgbiAKMDAwMDAwNjMxNCAwMDAwMCBuIAowMDAwMDA2OTA5IDAwMDAwIG4gCjAwMDAwMDY4ODYgMDAwMDAgbiAKMDAwMDAwNzUwNiAwMDAwMCBuIAowMDAwMDA3NDgzIDAwMDAwIG4gCjAwMDAwMDc2ODYgMDAwMDAgbiAKMDAwMDAwNzc5OCAwMDAwMCBuIAp0cmFpbGVyCjw8IC9TaXplIDMyCiAgIC9Sb290IDMxIDAgUgogICAvSW5mbyAzMCAwIFIKPj4Kc3RhcnR4cmVmCjc4NTEKJSVFT0YK"
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


class ROC_frame(wx.Frame):

    """
    Class used for creating frames other than the main one
    """

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title='READ-only R.O.C: ' + title,
            pos=wx.DefaultPosition,
            size=wx.Size(1166, 570),
            style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL,
        )
        self.SetIcon(OICC_LOGO.GetIcon())
        self.title = title
        self.panel = wx.Panel(self)
        self.theDB = wordy.DBpath

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        self.roctop = []

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.ROC2OPIN = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Record of Change to Operator Instructions',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.ROC2OPIN.Wrap(-1)
        self.ROC2OPIN.SetFont(wx.Font(
            14,
            74,
            90,
            90,
            False,
            'Arial',
        ))

        bSizer2.Add(self.ROC2OPIN, 0, wx.ALL, 5)

        self.m_staticline1 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL)
        bSizer2.Add(self.m_staticline1, 0, wx.EXPAND | wx.ALL, 5)

        self.stat_PEOI_num = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Instruction Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_PEOI_num.Wrap(-1)
        bSizer2.Add(self.stat_PEOI_num, 0, wx.ALL, 5)

        self.deftxt = []
        for n in range(0, 7):
            self.deftxt.append(wx.TextCtrl(
                self.panel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0,))

        self.m_textCtrl_PEOI = self.deftxt[0]
        self.m_textCtrl_PEOI.SetEditable(False)
        self.m_textCtrl_PEOI.SetBackgroundColour((195, 195, 195))
        bSizer2.Add(self.m_textCtrl_PEOI, 1, wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_cust_pn = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Customer and part Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_cust_pn.Wrap(-1)
        bSizer3.Add(self.stat_cust_pn, 0, wx.ALL, 5)

        self.m_textCtrl_cust_PN = self.deftxt[1]
        self.m_textCtrl_cust_PN.SetEditable(False)
        self.m_textCtrl_cust_PN.SetBackgroundColour((195, 195, 195))
        bSizer3.Add(self.m_textCtrl_cust_PN, 1, wx.ALL, 5)

        self.m_staticline2 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer3.Add(self.m_staticline2, 0, wx.EXPAND | wx.ALL, 5)

        self.stat_Pek_PN = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Pektron Part Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_Pek_PN.Wrap(-1)
        bSizer3.Add(self.stat_Pek_PN, 0, wx.ALL, 5)

        self.m_textCtrl_pekPN = self.deftxt[2]
        self.m_textCtrl_pekPN.SetEditable(False)
        self.m_textCtrl_pekPN.SetBackgroundColour((195, 195, 195))
        bSizer3.Add(self.m_textCtrl_pekPN, 1, wx.ALL, 5)

        bSizer1.Add(bSizer3, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_prod_desc = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Product Description',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_prod_desc.Wrap(-1)
        bSizer4.Add(self.stat_prod_desc, 0, wx.ALL, 5)

        self.m_textCtrl_prod_desc = self.deftxt[3]
        self.m_textCtrl_prod_desc.SetEditable(False)
        self.m_textCtrl_prod_desc.SetBackgroundColour((195, 195, 195))
        bSizer4.Add(self.m_textCtrl_prod_desc, 1, wx.ALL, 5)

        self.stat_stages = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Stage numbers',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_stages.Wrap(-1)
        bSizer4.Add(self.stat_stages, 0, wx.ALL, 5)

        self.m_textCtrl_stages = self.deftxt[4]
        self.m_textCtrl_stages.SetEditable(False)
        self.m_textCtrl_stages.SetBackgroundColour((195, 195, 195))
        bSizer4.Add(self.m_textCtrl_stages, 0, wx.ALL, 5)

        bSizer1.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_process_desc = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Operation / Process',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_process_desc.Wrap(-1)
        bSizer5.Add(self.stat_process_desc, 0, wx.ALL, 5)

        self.m_textCtrl_proc = self.deftxt[5]
        self.m_textCtrl_proc.SetEditable(False)
        self.m_textCtrl_proc.SetBackgroundColour((195, 195, 195))
        bSizer5.Add(self.m_textCtrl_proc, 1, wx.ALL, 5)

        self.obsstat = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Obsolete?',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.obsstat.Wrap(-1)
        bSizer5.Add(self.obsstat, 0, wx.ALL, 5)

        self.obs = self.deftxt[6]
        self.obs.SetEditable(False)
        self.obs.SetBackgroundColour((195, 195, 195))
        bSizer5.Add(self.obs, 0, wx.ALL, 5)

        bSizer1.Add(bSizer5, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline3 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer6.Add(self.m_staticline3, 1, wx.ALL | wx.EXPAND, 5)

        bSizer1.Add(bSizer6, 0, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        bSizer1.Add(bSizer7, 0, wx.EXPAND, 5)

        bSizer61 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline31 = wx.StaticLine(self.panel, wx.ID_ANY,
                                            wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer61.Add(self.m_staticline31, 0, wx.ALL | wx.EXPAND, 5)

        bSizer52 = wx.BoxSizer(wx.VERTICAL)
        bSizer52a = wx.BoxSizer(wx.HORIZONTAL)

        self.roc_prev = WX_Grid(self.panel, wx.ID_ANY,
                                wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid

        rowz = self.rowcalc()
        self.found_you = rowz[1]
        self.rows = rowz[0]
        self.roc_prev.CreateGrid(self.rows, 10)
        self.roc_prev.EnableGridLines(True)
        self.roc_prev.EnableDragGridSize(False)
        self.roc_prev.SetMargins(0, 0)

        # Columns

        self.roc_prev.SetColSize(0, 60)
        self.roc_prev.SetColSize(1, 274)
        self.roc_prev.SetColSize(2, 274)
        self.roc_prev.SetColSize(3, 60)
        self.roc_prev.SetColSize(4, 60)
        self.roc_prev.SetColSize(5, 90)
        self.roc_prev.SetColSize(6, 90)
        self.roc_prev.SetColSize(7, 90)
        self.roc_prev.SetColSize(8, 120)
        self.roc_prev.SetColSize(9, 120)
        self.roc_prev.EnableDragColMove(False)
        self.roc_prev.EnableDragColSize(True)
        self.roc_prev.SetColLabelSize(30)
        self.roc_prev.SetColLabelValue(0, u'ISSUE')
        self.roc_prev.SetColLabelValue(1, u'DETAILS')
        self.roc_prev.SetColLabelValue(2, u'REASON')
        self.roc_prev.SetColLabelValue(3, u'PGS')
        self.roc_prev.SetColLabelValue(4, u'COPIES')
        self.roc_prev.SetColLabelValue(5, u'IMPLMNT')
        self.roc_prev.SetColLabelValue(6, u'RCVD')
        self.roc_prev.SetColLabelValue(7, u'DATE')
        self.roc_prev.SetColLabelValue(8, u'FILE LOCATION')
        self.roc_prev.SetColLabelValue(9, u'link')
        self.roc_prev.SetColLabelAlignment(wx.ALIGN_CENTRE,
                                           wx.ALIGN_CENTRE)

        # Rows

        self.roc_prev.EnableDragRowSize(False)
        self.roc_prev.SetRowLabelSize(0)
        self.roc_prev.SetRowLabelAlignment(wx.ALIGN_CENTRE,
                                           wx.ALIGN_CENTRE)

        # Cell Defaults

        self.roc_prev.SetDefaultCellAlignment(wx.ALIGN_LEFT,
                                              wx.ALIGN_TOP)
        self.roc_prev.GetTargetWindow().SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.roc_prev.GetGridWindow().Bind(wx.EVT_MOTION,
                                           self.OnMouseMotion)

        bSizer52.Add(self.roc_prev, 1, wx.ALL | wx.EXPAND, 5)
        bSizer61.Add(bSizer52, 1, wx.EXPAND
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer1.Add(bSizer61, 1, wx.EXPAND
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.roc_prev.HideCol(9)

        self.panel.SetSizer(bSizer1)
        self.panel.Layout()
        self.roc_prev.Bind(WX_EVT_GRID_SELECT_CELL, self.onSingleSelect)

        self.panel.Centre(wx.BOTH)

        # should I move this in to the def?

        self.load_data()
        self.panel.Show()
        self.Show()

    def OnMouseMotion(self, event):
        self.roc_prev.GetTargetWindow().SetCursor(wx.Cursor(wx.CURSOR_HAND))
        event.Skip()

    def onSingleSelect(self, event):
        self.currentlySelectedCell = (event.GetRow(), 9)
        path_to_f = \
            self.roc_prev.GetCellValue(self.currentlySelectedCell)
        if os.path.isfile(path_to_f):
            SP_Popen(path_to_f, shell=True)
        event.Skip()

    def rowcalc(self):
        if not os.path.isfile(self.theDB):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return
        with open(self.theDB, 'rb') as f:
            self.roctop = PKL_load(f)
        found_you = -1
        found2 = -1
        found_you = findme(self.roctop, self.title)
        if found_you >= 0:
            if len(self.roctop[found_you]) > 0:
                found2 = foolproof_finder(self.roctop[found_you], 1)
        return [found2, found_you]

    def load_data(self):
        found_you = self.found_you
        if found_you >= 0:
            self.m_textCtrl_cust_PN.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                               ]['cstpn'], True))
            self.m_textCtrl_proc.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                            ]['stagedesc'], True))
            self.m_textCtrl_pekPN.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                             ]['pekpn'], True))
            self.m_textCtrl_prod_desc.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                                 ]['desc'], True))
            self.m_textCtrl_stages.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                              ]['stageno'], True))
            self.m_textCtrl_PEOI.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                            ]['opno'], True))
            self.obs.SetLabel(safestr(self.roctop[found_you][0]['details'].get('obsolete', "False"), True))

            if len(self.roctop[found_you]) > 0:
                acounter = self.rows
                xr = 0
                msg = wx.BusyInfo('Please wait, this may take a while (depends on the number of issue levels released).')
                while acounter >= 1:
                    if acounter > 0:
                        self.roc_prev.SetCellValue(xr, 0,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['rev'], True))
                        self.roc_prev.SetCellValue(xr, 1,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['doc'], True))
                        self.roc_prev.SetCellValue(xr, 2,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['rfc'], True))
                        self.roc_prev.SetCellValue(xr, 3,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['pages'], True))
                        self.roc_prev.SetCellValue(xr, 4,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['copies'], True))
                        self.roc_prev.SetCellValue(xr, 5,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['impl'], True))
                        self.roc_prev.SetCellValue(xr, 6,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['rcvd'], True))
                        self.roc_prev.SetCellValue(xr, 7,
                                                   safestr(self.roctop[found_you][acounter]['revs'
                                                                                            ]['date'], True))
                        pdf_name = self.title + '_' \
                            + safestr(self.roctop[found_you][acounter]['revs'
                                                                       ]['rev']) + '.pdf'
                        try:
                            (res1, res2) = search_file(wordy.rootpath,
                                                       pdf_name)
                        except TypeError:
                            (res1, res2) = ('missing', '')
                        self.roc_prev.SetCellValue(xr, 9, safestr(res2))
                        self.roc_prev.SetCellValue(xr, 8, safestr(res1))

                    xr += 1
                    acounter -= 1

    def __del__(self):
        wordy.ROROC_frame_number = 0


class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):

    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, wx.ID_ANY, style=wx.LC_REPORT |
                             wx.SUNKEN_BORDER)
        CheckListCtrlMixin.__init__(self)
        ListCtrlAutoWidthMixin.__init__(self)


class RcVr(wx.Dialog):

    def __init__(self, *args, **kw):
        super(RcVr, self).__init__(*args, **kw)
        self.SetIcon(OICC_LOGO_5.GetIcon())
        self.isopen = False
        self.SetSize((1000, -1))
        whichwindow = (args[1])
        if whichwindow == 0:
            self.finalised_results = [g for g in wordy.for_rcv_list]
        elif whichwindow == 1:
            self.finalised_results = [g for g in wordy.group_print_list]
        else:
            self.Close()
            self.Destroy()
            return

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        leftPanel = wx.Panel(panel)
        rightPanel = wx.Panel(panel, size=(570, -1))

        self.list = CheckListCtrl(rightPanel)
        self.list.InsertColumn(0, 'PEOI number', width=130)
        self.list.InsertColumn(1, 'Rev', width=50)
        self.list.InsertColumn(2, 'Description', width=260)
        self.list.InsertColumn(3, 'Details of Change', width=260)
        self.list.InsertColumn(4, 'PDF', width=4)
        self.list.InsertColumn(5, 'receiver', width=130)

        idx = 0
        self.receiverlist = []
        for i in self.finalised_results:
            self.receiverlist.append(i[5])
            index = self.list.InsertItem(idx, i[0])
            self.list.SetItem(index, 1, i[1])
            self.list.SetItem(index, 2, i[2])
            self.list.SetItem(index, 3, i[3])
            self.list.SetItem(index, 4, i[4])
            self.list.SetItem(index, 5, i[5])
            idx += 1
        self.receiverlist = list(set(self.receiverlist))

        vbox2 = wx.BoxSizer(wx.VERTICAL)

        config = configparser.ConfigParser()
        config.read('settings.ini')
        try:
            self.me = (safestr(config['user_profile']['me']))
        except KeyError:
            self.me = "UNSET"
        signedbyme = " Files will be marked as \n signed by {}".format(self.me)
        selBtn = wx.Button(leftPanel, label='Select All')
        desBtn = wx.Button(leftPanel, label='Deselect All')
        selByBtn = wx.Button(leftPanel, label='Select By')
        self.recvrscombos = wx.ComboBox(
            leftPanel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
        )
        for m in self.receiverlist:
            self.recvrscombos.Append(m)
        if whichwindow == 0:
            appBtn = wx.Button(leftPanel, label='Mark Received')
            note = wx.StaticText(leftPanel, label=signedbyme)
            self.Bind(wx.EVT_BUTTON, self.GroupSignFor, id=appBtn.GetId())
        elif whichwindow == 1:
            appBtn = wx.Button(leftPanel, label='Group Open')
            note = wx.StaticText(leftPanel, label="")
            self.Bind(wx.EVT_BUTTON, self.GroupOpen, id=appBtn.GetId())
            appBtn2 = wx.Button(leftPanel, label='txt')
            self.Bind(wx.EVT_BUTTON, self.Checkup, id=appBtn2.GetId())

        self.Bind(wx.EVT_BUTTON, self.OnSelectAll, id=selBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnDeselectAll, id=desBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnSelBy, id=selByBtn.GetId())

        vbox2.Add(selBtn, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(self.recvrscombos, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(selByBtn, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(desBtn, 3, wx.EXPAND, 1)
        # Create a horizontal box sizer
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add the two buttons to the horizontal sizer
        btn_sizer.Add(appBtn, 1, wx.EXPAND, 0)  # Proportion is set to 1, so this button will be resized
        if whichwindow == 1:
            btn_sizer.Add(appBtn2, 0, wx.EXPAND, 0)  # Proportion is set to 0, so this button will not be resized and will take up only the space it needs

        # Add the horizontal sizer to the main vertical sizer (vbox2)
        vbox2.Add(btn_sizer, 3, wx.EXPAND, 1)

        vbox2.Add(note, 3, wx.EXPAND, 5)

        leftPanel.SetSizer(vbox2)

        vbox.Add(self.list, 4, wx.EXPAND | wx.TOP, 3)
        vbox.Add((-1, 10))

        rightPanel.SetSizer(vbox)

        hbox.Add(leftPanel, 0, wx.EXPAND | wx.RIGHT, 5)
        hbox.Add(rightPanel, 1, wx.EXPAND)
        hbox.Add((3, -1))

        panel.SetSizer(hbox)
        if whichwindow == 0:
            self.SetTitle('Mark as Received')
        elif whichwindow == 1:
            self.SetTitle('Group Open for Print')
        self.Centre()

    def OnSelectAll(self, event):
        self.OnDeselectAll(None)
        num = self.list.GetItemCount()
        for i in range(num):
            self.list.CheckItem(i)

    def OnSelBy(self, event):
        self.OnDeselectAll(None)
        if len(self.recvrscombos.GetValue()) > 0:
            num = self.list.GetItemCount()
            for i in range(num):
                if self.recvrscombos.GetValue() == self.list.GetItemText(i, 5):
                    self.list.CheckItem(i)

    def OnDeselectAll(self, event):

        num = self.list.GetItemCount()
        for i in range(num):
            self.list.CheckItem(i, False)

    def GroupSignFor(self, event):
        self.theDB = wordy.DBpath
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            print("DB error")
            return
        num = self.list.GetItemCount()
        for i in range(num):
            if self.list.IsChecked(i):
                self.add_to_DB(self.list.GetItemText(i, 0),
                               self.list.GetItemText(i, 1))
                msg = wx.BusyInfo(
                    'marking %s rev %s as received by %s' % (self.list.GetItemText(i, 0), self.list.GetItemText(i, 1), self.me))
                sleep(0.20)
        with open(self.theDB, 'wb') as f:
            PKL_dump(self.roctop, f, indent=2)
        print(datetime.datetime.now())
        msg = wx.BusyInfo('Complete!')
        sleep(0.4)
        self.Close()

    def add_to_DB(self, toof, threef):
        kill = "z"
        killer = "z"
        xcount = 0
        for x in self.roctop:
            if x[0]['details']['opno'] == toof:
                kill = xcount
                break
            xcount += 1
        if kill == "z":
            print("PEOI number %s not found." % (toof))
            return
        if kill != "z":
            c = 0
            for z in self.roctop[kill]:
                try:
                    if z['revs']['rev'] == threef:
                        killer = c
                except KeyError:
                    pass
                c += 1
            if killer == "z":
                print ('Revision not there')
                return
            elif killer != "z":
                self.roctop[kill][killer]['revs']['date'] = safestr(
                    datetime.datetime.today().strftime('%d-%m-%Y'))
                self.roctop[kill][killer]['revs']['rcvd'] = self.me

    def shorten(self, string2short, len2short):
        if len(string2short) > len2short:
            return str(string2short[0:len2short-1] + "...")
        else:
            return string2short

    def Checkup(self, event):
        self.GroupOpen(None, called_from_checkup=True)
        self.Close()
        return

    def GroupOpen(self, event, called_from_checkup=False):
        num = self.list.GetItemCount()
        total = 0
        for i in range(num):
            if self.list.IsChecked(i):
                total += 1
        if total == 0:
            return
        csvdetails = ["PEOI NUMBER", "REV", "DESCRIPTION", "DETAILS OF CHANGE"]
        csvrows = []
        for i in range(num):
            if self.list.IsChecked(i):
                templist = []
                x = self.list.GetItemText(i, 4)
                self.manager = self.list.GetItemText(i, 5)
                x1, x2, x3, x4 = self.list.GetItemText(i, 0).strip(), self.list.GetItemText(i, 1).strip(), self.list.GetItemText(i, 2).strip(), self.list.GetItemText(i, 3).strip()
                templist.append(self.shorten(x1, 25))
                templist.append(self.shorten(x2, 25))
                templist.append(self.shorten(x3, 45))
                templist.append(self.shorten(x4, 95))
                csvrows.append(templist)
                fpath = x[:len(x) - 4]
                out_file = os.path.abspath(wordy.forappralpath + '\\'
                                           + fpath + '\\' + x)
                if not called_from_checkup:
                    if os.path.isfile(out_file):
                        SP_Popen([out_file], shell=True)
        col1 = []
        col2 = []
        col3 = []
        col4 = []
        longs = []
        for rows in csvrows:
            col1.append(rows[0])
            col2.append(rows[1])
            col3.append(rows[2])
            col4.append(rows[3])
        cols = [col1, col2, col3, col4]
        for col in cols:
            try:
                longest = max(col, key=len)
                longs.append(len(longest))
            except ValueError:
                longs.append(0)
        for rows in csvrows:
            rows[0] = rows[0].ljust(16)
            rows[1] = rows[1].ljust(8)
            rows[2] = rows[2].ljust(longs[2]+4)
            rows[3] = rows[3].ljust(longs[3]+4)
        csvdetails[0] = csvdetails[0].ljust(16)
        csvdetails[1] = csvdetails[1].ljust(8)
        csvdetails[2] = csvdetails[2].ljust(longs[2]+4)
        csvdetails[3] = csvdetails[3].ljust(longs[3]+4)
        try:
            tf1 = os.path.abspath(wordy.temp_directory + "\\" + 'justprinted.txt')
            tf2 = os.path.abspath(wordy.temp_directory + "\\" + 'justprinted.pdf')
            with open(tf1, 'w') as f:
                write = csv.writer(f, delimiter='\t', doublequote=False, escapechar='\\', quoting=csv.QUOTE_NONE)
                write.writerow(["            OPERATOR INSTRUCTIONS TOP SHEET.     DATE: ", safestr(datetime.datetime.today().strftime('%d-%m-%Y'))])
                write.writerow([self.manager, "- you have been issued with the following (new or updated) laminated instructions:"])
                write.writerow(csvdetails)
                write.writerows(csvrows)
                write.writerow(["Please mark these as received in the OICC software (Finalise > Open Receiver Window)."])
            if called_from_checkup:
                if os.path.isfile(tf1):
                    SP_Popen([tf1], shell=True)
                return
            createpdf = FPDF()
            createpdf.add_page(orientation='L')
            createpdf.set_font("Courier", size=10)
            text_file = open(tf1, "r")
            for text in text_file:
                createpdf.cell(220, 12, txt=text, ln=1, align='L')
            createpdf.output(tf2)
            if os.path.isfile(tf2):
                SP_Popen([tf2], shell=True)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            print("temp file / PDF gen error")
        self.Close()


class FinaliseFrame(wx.Frame):

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title=title,
            pos=wx.DefaultPosition,
            size=wx.Size(1000, 500),
            style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU |
            wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN | wx.TAB_TRAVERSAL,
        )
        
        self.SetIcon(OICC_LOGO_4.GetIcon())
        self.theDB = wordy.DBpath
        leftsizer = wx.BoxSizer(wx.VERTICAL)
        rightsizer = wx.BoxSizer(wx.VERTICAL)
        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)
        bSizer3.Add(leftsizer, proportion=1, flag=wx.EXPAND)
        bSizer3.Add(rightsizer, proportion=1, flag=wx.EXPAND)
        fontdefault = wx.Font(
            10,
            75,
            90,
            92,
            False,
            'Consolas',
        )

        bSizer10 = wx.BoxSizer(wx.VERTICAL)
        self.sign_m_staticText2 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'For Print',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.sign_m_static_markrcvd = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Sign as Received',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.sign_m_static_markrcvd.Wrap(-1)
        self.sign_m_staticText2.Wrap(-1)
        self.sign_m_staticText2.SetFont(fontdefault)
        self.sign_m_static_markrcvd.SetFont(fontdefault)
        rightbSizer10 = wx.BoxSizer(wx.VERTICAL)
        self.rightsign_m_staticText2 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Finalisation',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.rightsign_m_staticText2.Wrap(-1)
        self.rightsign_m_staticText2.SetFont(fontdefault)
        bSizer10.Add(self.sign_m_static_markrcvd, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.open_mark_rcvd = wx.Button(
            self,
            wx.ID_ANY,
            u'Open Receiver Window',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.open_mark_rcvd.Disable()
        if wordy.userlevel == "PD" or wordy.userlevel == "master":
            self.open_mark_rcvd.Enable()

        bSizer10.AddSpacer(4)

        bSizer10.Add(self.open_mark_rcvd, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.group_opener = wx.Button(
            self,
            wx.ID_ANY,
            u'Group open for Print',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        bSizer10.AddSpacer(4)
        bSizer10.Add(self.sign_m_staticText2, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer10.Add(self.group_opener, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)
        rightbSizer10.Add(self.rightsign_m_staticText2, 0, wx.ALL
                          | wx.ALIGN_CENTER_HORIZONTAL, 5)
        leftsizer.Add(bSizer10, proportion=1, flag=wx.ALIGN_CENTER)
        rightsizer.Add(rightbSizer10, proportion=1, flag=wx.ALIGN_CENTER)
        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)
        bSizer4.AddSpacer(10)
        rightbSizer4 = wx.BoxSizer(wx.HORIZONTAL)
        rightbSizer4.AddSpacer(10)
        self.generatecombolists()

        self.m_comboBox2 = wx.ComboBox(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            self.m_comboBox2Choices,
            0,
        )
        bSizer4.Add(self.m_comboBox2, 2, wx.ALL, 5)

        self.sign_m_button2 = wx.Button(
            self,
            wx.ID_ANY,
            u'open (to print)',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer4.Add(self.sign_m_button2, 2, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer4.AddSpacer(10)

        self.rightm_comboBox2 = wx.ComboBox(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            self.rightm_comboBox2Choices,
            0,
        )
        rightbSizer4.Add(self.rightm_comboBox2, 2, wx.ALL, 5)

        rightbSizer4.AddSpacer(5)

        leftsizer.Add(bSizer4, 0, wx.EXPAND, 5)

        self.sign_m_button3 = wx.Button(
            self,
            wx.ID_ANY,
            u'Finalise',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.rightsign_m_button3 = wx.Button(
            self,
            wx.ID_ANY,
            u'Execute Group Finalisation',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        rightbSizer10.Add(self.rightsign_m_button3, 5, wx.ALL
                          | wx.ALIGN_CENTER_HORIZONTAL, 5)
        rightbSizer4.Add(self.sign_m_button3, 5, wx.ALL
                         | wx.ALIGN_CENTER_HORIZONTAL, 5)
        rightsizer.Add(rightbSizer4, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.HORIZONTAL)
        bSizer6.AddSpacer(50)
        self.m_staticline1 = wx.StaticLine(self, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer6.Add(self.m_staticline1, 0, wx.EXPAND | wx.ALL, 5)

        leftsizer.Add(bSizer6, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_comboBox2.Bind(wx.EVT_COMBOBOX,
                              self.m_comboBox2OnCombobox2)
        self.rightm_comboBox2.Bind(wx.EVT_COMBOBOX,
                                   self.m_comboBox2OnCombobox)
        self.sign_m_button2.Bind(wx.EVT_BUTTON, self.viewer)
        self.sign_m_button3.Bind(wx.EVT_BUTTON, self.finalise)
        self.rightsign_m_button3.Bind(wx.EVT_BUTTON, self.group_finalise)
        self.open_mark_rcvd.Bind(wx.EVT_BUTTON, self.mark_as_rcvd)
        self.group_opener.Bind(wx.EVT_BUTTON, self.do_grp_opn)
        self.fname = self.m_comboBox2.GetValue()
        self.master = {
            'PE_sign': 'Prod.Eng sign-off',
            'PD_sign': 'Prod.Dept sign-off',
            'QA_sign': 'QA.Dept sign-off',
            'doc': 'Details of Change',
            'rev': 'Revision',
            'copies': 'number of copies issued',
            'rfc': 'Reason for Change',
            'date': 'date received by Production',
            'rcvd': 'received for production by',
            'pages': 'number of pages',
            'impl': 'implemented by',
        }
        self.m_textCtrl11 = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 30),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )

        rightbSizer6 = wx.BoxSizer(wx.HORIZONTAL)
        rightbSizer6.AddSpacer(100)
        self.rightm_staticline1 = wx.StaticLine(self, wx.ID_ANY,
                                                wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        rightbSizer6.Add(self.rightm_staticline1, 0, wx.EXPAND | wx.ALL, 5)
        rightsizer.Add(bSizer6, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.rightm_textCtrl11 = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 30),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )

        leftsizer.Add(self.m_textCtrl11, 2, wx.EXPAND, 5)
        rightsizer.Add(self.rightm_textCtrl11, 2, wx.EXPAND, 5)
        self.SetSizer(bSizer3)
        self.DOC = ''
        self.RFC = ''

        try:
            self.m_textCtrl11.Clear()
            for prx in self.printreport:
                self.m_textCtrl11.WriteText(prx + '\n')
            self.rightm_textCtrl11.Clear()
            for prx in self.printreport2:
                self.rightm_textCtrl11.WriteText(prx + '\n')
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        self.Layout()
        if wordy.reportfunc == 0:
            self.Show()
        elif wordy.reportfunc == 1:
            out_file = os.path.abspath(wordy.forappralpath + '\\'
                                       + 'report.txt')
            with open(out_file, 'w+') as f:
                for xx in self.m_comboBox2ChoicesFULL:
                    (pp, result) = self.load_data(xx)
                    for z in pp:
                        f.write(z)
                        f.write('\n')
                    f.write('\n')
                    f.write('------------')
                    f.write('\n')
            SP_Popen([out_file], shell=True)
            wordy.reportfunc = 0
            self.Close()
            return None
        elif wordy.reportfunc == 2:
            wordy.reportfunc = 0
            self.Close()
            return None

    def generatecombolists(self):
        self.m_comboBox2ChoicesFULL = self.getfile_list()
        self.m_comboBox2Choices = \
            self.filterIT(self.m_comboBox2ChoicesFULL)
        self.rightm_comboBox2Choices = \
            self.strict_filterIT(self.m_comboBox2ChoicesFULL)

    def getfile_list(self):
        combo = []
        for (dirpath, dirnames, fi) in os.walk(wordy.forappralpath):
            for fp in fi:
                if fp.endswith('.db'):
                    try:
                        killitdead = os.path.abspath(dirpath + "\\" + fp)
                        print(killitdead)
                        os.remove(killitdead)
                    except Exception as e:
                        print('oops ' + str((inspect.stack()[0][2])))
                        print (e.message, e.args)
                        pass
                if not fp.endswith(('.txt', '.db')):
                    combo.append(fp)
            if fi:
                pass
            if not fi:
                try:
                    if dirpath != wordy.forappralpath:
                        os.rmdir(dirpath)
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    pass
        return combo

    def mark_as_rcvd(self, event):
        try:
            self.ex.Close()
        except AttributeError:
            pass
        except RuntimeError:
            pass
        passiton = []
        self.forrcvd = \
            self.reverse_strict_filterIT(self.m_comboBox2ChoicesFULL)
        for x in self.forrcvd:
            xx = self.load_good_data(x)
            passiton.append(xx)
        wordy.for_rcv_list = passiton
        self.ex = RcVr(None, 0)
        if not self.ex.isopen:
            self.ex.isopen = True
            self.ex.ShowModal()
        self.updatecombos()

    def updatecombos(self):
        self.generatecombolists()
        self.m_comboBox2.Clear()
        self.m_comboBox2.AppendItems(self.m_comboBox2Choices)
        self.rightm_comboBox2.Clear()
        self.rightm_comboBox2.AppendItems(self.rightm_comboBox2Choices)
        try:
            self.m_textCtrl11.Clear()
            for prx in self.printreport:
                self.m_textCtrl11.WriteText(prx + '\n')
            self.rightm_textCtrl11.Clear()
            for prx in self.printreport2:
                self.rightm_textCtrl11.WriteText(prx + '\n')
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass

    def do_grp_opn(self, event):
        try:
            self.grp.Close()
        except AttributeError:
            pass
        except RuntimeError:
            pass
        passiton = []
        self.forrcvd = self.reverse_strict_filterIT(
            self.m_comboBox2ChoicesFULL)
        for x in self.forrcvd:
            xx = self.load_good_data(x)
            passiton.append(xx)
        wordy.group_print_list = passiton
        self.grp = RcVr(None, 1)
        if not self.grp.isopen:
            self.grp.isopen = True
            self.grp.ShowModal()
        self.updatecombos()

    def chKr(self, from_viewer):
        if from_viewer == 1:
            self.fname = self.m_comboBox2.GetValue()
        elif from_viewer == 2:
            self.fname = self.rightm_comboBox2.GetValue()
        else:
            self.fname = from_viewer
        if self.fname != '':
            fpath = self.fname[:len(self.fname) - 4]
            out_file = os.path.abspath(wordy.forappralpath + '\\'
                                       + fpath + '\\' + self.fname)
            if os.path.isfile(out_file):
                return out_file
            else:
                return 0
        else:
            return 0

    def viewer(self, event):
        n1 = self.chKr(1)
        if n1 != 0:
            SP_Popen([n1], shell=True)
        return

    def log_and_store(self, log_item):
        logger.info(log_item)
        self.reportlist.append(log_item)
        return

    def group_finalise(self, event):
        for x in self.rightm_comboBox2Choices:
            self.finalise(None, self.chKr(x))
        print("group sign off complete")
        wordy.suppress = 0
        self.Close()
        x = FinaliseFrame('Finaliser (re-opened)')

    def finalise(self, event, do_group=False):
        self.reportlist = []
        go2obs = None
        if do_group:
            n1 = do_group
        else:
            n1 = self.chKr(2)
        if n1 != 0:
            (reePort, result) = self.load_data(self.fname)
            try:
                cr1 = self.current_record[0]
                cr2 = self.current_record[1]
                revm1 = cr1[cr2 - 1].get('revs', {}).get('rev')
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                revm1 = False
            if result:
                wordy.suppress = 1
                proj = opno2proj(self.fname)
                if proj != '':
                    appdir = os.path.abspath(wordy.approvedpath + '\\'
                                             + proj)
                    wordy(None, 'nowt').makepaths(appdir)
                    go_to = os.path.abspath(appdir + '\\' + self.fname)
                    (REV, PEOI) = getRnP(self.fname)
                    go_from = os.path.abspath(wordy.forappralpath + '\\'
                                              + PEOI + '_' + REV + '\\' + self.fname)
                    gofdir = os.path.abspath(wordy.forappralpath + '\\'
                                             + PEOI + '_' + REV)
                    if revm1:
                        go2obs = os.path.abspath(
                            appdir + '\\' + PEOI + '_' + revm1 + ".pdf")
                        if not os.path.isfile(go2obs):
                            go2obs = None
                    if os.path.isfile(go_from):
                        if not os.path.isfile(go_to):
                            if os.path.exists(appdir):
                                msg = wx.BusyInfo(
                                    'moving %s to %s (approved directory)' % (self.fname, appdir))
                                try:
                                    if fileaccesscheck(go_from):
                                        if not fileaccesscheck(go_to):
                                            os.rename(go_from, go_to)

                                            sleep(0.5)
                                            msg = wx.BusyInfo('done')
                                            sleep(0.25)
                                            try:
                                                run_once()
                                                self.log_and_store(
                                                    ctime(time()))
                                                self.log_and_store(self.fname)
                                                self.log_and_store(go_to)
                                                self.log_and_store(
                                                    'details of change: ' + str(self.DOC))
                                                self.log_and_store(
                                                    'reasons for change: ' + str(self.RFC))
                                            except Exception as e:
                                                print('oops ' + str((inspect.stack()[0][2])))
                                                print (e.message, e.args)
                                                pass
                                except IOError:
                                    dlg = wx.MessageDialog(None,
                                                           'File in use, cannot move. Please close the PDF first!', '',
                                                           wx.OK | wx.ICON_ERROR)
                                    dlg.ShowModal()
                                    dlg.Destroy()
                                if go2obs is not None:
                                    if (wx.MessageBox((PEOI + '_' + revm1 + " found in approved directory. This is probably the previous revision. Would you like to make this obsolete now?"), ("make obsolete option"), parent=self, style=wx.YES_NO | wx.ICON_WARNING,) == wx.NO):
                                        pass
                                    else:
                                        wordy.m_filePicker1 = go2obs
                                        wordy(None, title='duplicrap').m_filePicker1OnFileChanged(
                                            wordy)
                                        wordy.m_filePicker1 = None
                                        self.log_and_store(
                                            'previous revision, ' + PEOI + '_' + revm1 + ' made obsolete')
                                try:
                                    if gofdir != wordy.forappralpath:
                                        os.rmdir(gofdir)
                                except Exception as e:
                                    print('oops ' + str((inspect.stack()[0][2])))
                                    print (e.message, e.args)
                                    pass
                                logger.info('-----')
                                logger.info('')
                                connection_obj = sqlite3.connect(
                                    wordy.loggingDB)
                                # cursor object
                                cursor_obj = connection_obj.cursor()
                                # Creating table
                                table = """ CREATE TABLE IF NOT EXISTS updated_instructions (
                                            Time VARCHAR(255) NOT NULL,
                                            PEOI VARCHAR(255) NOT NULL,
                                            Link VARCHAR(255) NOT NULL,
                                            Details VARCHAR(255) NOT NULL,
                                            Reasons VARCHAR(255) NOT NULL
                                        ); """
                                cursor_obj.execute(table)
                                connection_obj.execute(
                                    'INSERT INTO updated_instructions VALUES (?,?,?,?,?)', self.reportlist[:5])
                                connection_obj.commit()
                                # Close the connection
                                connection_obj.close()
                if do_group:
                    return
                wordy.suppress = 0
                self.Close()
                x = FinaliseFrame('Finaliser (re-opened)')

        return

    def roguekill(self, PEOI, REV, PEOIx):
        nullfile = os.path.abspath(wordy.forappralpath + '\\' + PEOI + '_' + REV + '\\' + PEOIx)
        nullpath = os.path.abspath(wordy.forappralpath + '\\' + PEOI + '_' + REV)
        if os.path.isfile(nullfile):
            try:
                if fileaccesscheck(nullfile):
                    os.remove(nullfile)
                    if nullpath != wordy.forappralpath:
                        print ("error: " + PEOIx + " is not in database - file and folder will be removed")
                        os.rmdir(nullpath)
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                pass

    def filterIT(self, lizt):
        lizzt = []
        prntlzt = []
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        for PEOIx in lizt:
            (REV, PEOI) = getRnP(PEOIx)
            found_you = findme(self.roctop, PEOI)
            if found_you >= 0:
                found2 = moreFinding(self.roctop[found_you], REV)
                if found2 > 0:
                    try:
                        if self.roctop[found_you][found2]['revs']['PE_sign'
                                                                ] == 1:
                            if self.roctop[found_you][found2]['revs'
                                                            ]['PD_sign'] == 1:
                                if self.roctop[found_you][found2]['revs'
                                                                ]['QA_sign'] == 1:
                                    if self.roctop[found_you][found2]['revs']['date'] == '' or self.roctop[found_you][found2]['revs']['rcvd'] == '':
                                        lizzt.append(PEOIx)
                                        templizt = [
                                            PEOIx, self.roctop[found_you][found2]['revs']['copies']]
                                        prntlzt.append(templizt)
                                        rej_file = \
                                            os.path.abspath(wordy.forappralpath
                                                            + '\\' + PEOI + '_' + REV + '\\'
                                                            + PEOI + '_' + REV + '.txt')
                                        if os.path.isfile(rej_file):
                                            os.remove(rej_file)
                    except:
                        self.roguekill(PEOI, REV, PEOIx)
            else:
                self.roguekill(PEOI, REV, PEOIx)

        self.printreport = []
        self.printreport.append(
            'Below are Operator Instructions that are signed by all departments.')
        self.printreport.append(
            'These need printing out & laminating & distributing to managers.')
        self.printreport.append(
            'They can then be marked as received, and finalised in the software.')
        self.printreport.append(
            ' Note: the open to print button is disabled if there are empty fields.')
        for prx in prntlzt:
            linew = str(prx[0]) + ". # Copies to print: " + str(prx[1])
            self.printreport.append(str(linew))
        return lizzt

    def strict_filterIT(self, lizt):
        lizzt = []
        prntlzt = []
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        for PEOIx in lizt:
            (REV, PEOI) = getRnP(PEOIx)
            found_you = findme(self.roctop, PEOI)
            if found_you >= 0:
                found2 = moreFinding(self.roctop[found_you], REV)
                if found2 > 0:
                    try:
                        if self.roctop[found_you][found2]['revs']['PE_sign'] == 1:
                            if self.roctop[found_you][found2]['revs']['PD_sign'] == 1:
                                if self.roctop[found_you][found2]['revs']['QA_sign'] == 1:
                                    if self.roctop[found_you][found2]['revs']['copies'] != '':
                                        if self.roctop[found_you][found2]['revs']['date'] != '':
                                            if self.roctop[found_you][found2]['revs']['rcvd'] != '':
                                                if self.roctop[found_you][found2]['revs']['pages'] != '':
                                                    if self.roctop[found_you][found2]['revs']['impl'] != '':
                                                        lizzt.append(PEOIx)
                                                        templizt = [
                                                            PEOIx, self.roctop[found_you][found2]['revs']['copies']]
                                                        prntlzt.append(templizt)
                                                        rej_file = \
                                                            os.path.abspath(wordy.forappralpath
                                                                            + '\\' + PEOI + '_' + REV + '\\'
                                                                            + PEOI + '_' + REV + '.txt')
                                                        if os.path.isfile(rej_file):
                                                            os.remove(rej_file)
                    except IndexError:
                        self.roguekill(PEOI, REV, PEOIx)
            else:
                self.roguekill(PEOI, REV, PEOIx)
        self.printreport2 = []
        self.printreport2.append(
            'Below are Operator Instructions that are ready to be Finalised. They have been received')
        self.printreport2.append(
            'by the Manager(s). You can finalise these one at a time, or use the Group Finalise')
        self.printreport2.append(
            'Tool to do them all. Finalising moves the PDF to "approved" & removes PEOI from WIP')
        for prx in prntlzt:
            linew = str(prx[0])
            self.printreport2.append(str(linew))
        return lizzt

    def reverse_strict_filterIT(self, lizt):
        lizzt = []
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        for PEOIx in lizt:
            (REV, PEOI) = getRnP(PEOIx)
            found_you = findme(self.roctop, PEOI)
            if found_you >= 0:
                found2 = moreFinding(self.roctop[found_you], REV)
                if found2 > 0:
                    if self.roctop[found_you][found2]['revs']['PE_sign'] == 1:
                        if self.roctop[found_you][found2]['revs']['PD_sign'] == 1:
                            if self.roctop[found_you][found2]['revs']['QA_sign'] == 1:
                                if self.roctop[found_you][found2]['revs']['copies'] != '':
                                    if self.roctop[found_you][found2]['revs']['pages'] != '':
                                        if self.roctop[found_you][found2]['revs']['impl'] != '':
                                            if self.roctop[found_you][found2]['revs']['date'] == '' or self.roctop[found_you][found2]['revs']['rcvd'] == '':
                                                lizzt.append(PEOIx)
        return lizzt

    def load_data(self, PEOIx):
        OKGO = False
        report = []
        self.good_lizt = []
        self.bad_lizt = []
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        report.append(PEOIx)
        found_you = -1
        found2 = -1
        (REV, PEOI) = getRnP(PEOIx)
        found_you = findme(self.roctop, PEOI)
        if found_you >= 0:
            found2 = moreFinding(self.roctop[found_you], REV)
            if found2 > 0:
                self.DOC = safestr(
                    self.roctop[found_you][found2]['revs']['doc'])
                self.RFC = safestr(
                    self.roctop[found_you][found2]['revs']['rfc'])
                if self.roctop[found_you][found2]['revs']['PE_sign'] \
                        == 1:
                    self.good_lizt.append('PE_sign')
                else:
                    self.bad_lizt.append('PE_sign')
                if self.roctop[found_you][found2]['revs']['PD_sign'] \
                        == 1:
                    self.good_lizt.append('PD_sign')
                else:
                    self.bad_lizt.append('PD_sign')
                if self.roctop[found_you][found2]['revs']['QA_sign'] \
                        == 1:
                    self.good_lizt.append('QA_sign')
                else:
                    self.bad_lizt.append('QA_sign')
                for oknok in self.roctop[found_you][found2]['revs']:
                    if safestr(oknok) != 'PE_sign' and safestr(oknok) \
                            != 'PD_sign' and safestr(oknok) != 'QA_sign':
                        if self.roctop[found_you][found2]['revs'
                                                          ][oknok] != '':
                            self.good_lizt.append(oknok)
                        else:
                            self.bad_lizt.append(oknok)
            elif found2 == -1:
                report.append('no record of change for this instruction at this revision level'
                              )
        elif found_you == -1:
            report.append('no record of change for this instruction at any revision level'
                          )

        # use self.master dict and goodbad lists to display what is Ok and what is not OK

        xl = 1
        OKGO = True
        for x in self.master:
            for y in self.good_lizt:
                if x == y:
                    xl += 1
            for p in self.bad_lizt:
                if x == p:
                    if 'received' not in self.master[p]:
                        report.append(self.master[p] + '  - needs completing')
                        OKGO = False
                    else:
                        report.append(
                            self.master[p] + '  - when printed & laminted, \nneeds marking as received')
                    xl += 1

        if OKGO:
            report.append(PEOIx
                          + ' is ready to be printed'
                          )
            self.current_record = [self.roctop[found_you], found2]
        return (report, OKGO)

    def load_good_data(self, PEOIx):
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        (REV, PEOI) = getRnP(PEOIx)
        found_you = findme(self.roctop, PEOI)
        if found_you >= 0:
            found2 = moreFinding(self.roctop[found_you], REV)
            if found2 > 0:
                DTLS = safestr(
                    self.roctop[found_you][0]['details']['stagedesc'])
                DOC = safestr(
                    self.roctop[found_you][found2]['revs']['doc'])
                RCVR = safestr(
                    self.roctop[found_you][found2]['revs']['PD_sign_name'])

        return (PEOI, REV, DTLS, DOC, PEOIx, RCVR)

    def m_comboBox2OnCombobox(self, event):
        fname = self.rightm_comboBox2.GetValue()
        if fname != '':
            if fname != self.fname:
                self.fname = fname
                (reePort, result) = self.load_data(self.fname)
                self.rightm_textCtrl11.Clear()
                for z in reePort:
                    zz = z.replace("is ready to be printed", "is ready to be FINALISED")
                    self.rightm_textCtrl11.WriteText(zz + '\n')
                if not result:
                    self.sign_m_button3.Disable()
                    self.sign_m_button3.SetLabel('--------------------')
                elif result:
                    self.sign_m_button3.Enable()
                    self.sign_m_button3.SetLabel('Execute Finalisation')

    def m_comboBox2OnCombobox2(self, event):
        fname = self.m_comboBox2.GetValue()
        if fname != '':
            if fname != self.fname:
                self.fname = fname
                (reePort, result) = self.load_data(self.fname)
                self.m_textCtrl11.Clear()
                for z in reePort:
                    self.m_textCtrl11.WriteText(z + '\n')
                if not result:
                    self.sign_m_button2.Disable()
                    self.sign_m_button2.SetLabel('complete outstanding')
                elif result:
                    self.sign_m_button2.Enable()
                    self.sign_m_button2.SetLabel('Open and Print')

    def xferttech(self, Xto, Xfrom):
        try:
            distutils.dir_util.remove_tree(Xto)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        try:
            distutils.dir_util.copy_tree(Xfrom, Xto)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        return

    def __del__(self):
        wordy.final_frame_number = 0


class ApproveFrame(wx.Frame):

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title=title,
            pos=wx.DefaultPosition,
            size=wx.Size(800, 500),
            style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU |
            wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN | wx.TAB_TRAVERSAL,
        )
        
        self.SetIcon(OICC_LOGO_3.GetIcon())
        self.theDB = wordy.DBpath

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        bSizer10 = wx.BoxSizer(wx.VERTICAL)
        header = wordy.origin \
            + u" Digital Sign off of PEOI from 'awaiting approval' directory"

        self.sign_m_staticText2 = wx.StaticText(
            self,
            wx.ID_ANY,
            header,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.sign_m_staticText2.Wrap(-1)
        self.sign_m_staticText2.SetFont(wx.Font(
            10,
            75,
            90,
            92,
            False,
            'Consolas',
        ))

        bSizer10.Add(self.sign_m_staticText2, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer3.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer4.AddSpacer(10)

        self.m_comboBox2ChoicesFULL = []
        for (dirpath, dirnames, fi) in os.walk(wordy.forappralpath):
            for fp in fi:
                if fp.endswith('.db'):
                    try:
                        killitdead = os.path.abspath(dirpath + "\\" + fp)
                        print(killitdead)
                        os.remove(killitdead)
                    except Exception as e:
                        print('oops ' + str((inspect.stack()[0][2])))
                        print (e.message, e.args)
                        pass
                if not fp.endswith(('.txt', '.db')):
                    self.m_comboBox2ChoicesFULL.append(fp)

        self.m_comboBox2Choices = \
            self.filterIT(self.m_comboBox2ChoicesFULL)

        self.m_comboBox2 = wx.ComboBox(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            self.m_comboBox2Choices,
            0,
        )
        bSizer4.Add(self.m_comboBox2, 3, wx.ALL, 5)

        self.sign_m_button2 = wx.Button(
            self,
            wx.ID_ANY,
            u'open file in Cute PDF',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.sign_m_button2AR = wx.Button(
            self,
            wx.ID_ANY,
            u'Open in Adobe',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.fillfields = wx.Button(
            self,
            wx.ID_ANY,
            u'Auto',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        bSizer4.Add(self.sign_m_button2, 2, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer4.Add(self.fillfields, 1, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer4.Add(self.sign_m_button2AR, 2, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer4.AddSpacer(10)

        bSizer3.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer1011 = wx.BoxSizer(wx.VERTICAL)

        self.sign_m_staticText211 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Review the document and sign off if OK',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.sign_m_staticText211.Wrap(-1)
        self.sign_m_staticText211.SetFont(wx.Font(
            10,
            75,
            90,
            92,
            False,
            'Consolas',
        ))

        bSizer1011.Add(self.sign_m_staticText211, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer3.Add(bSizer1011, 0, wx.EXPAND, 5)

        bSizerx17 = wx.BoxSizer(wx.HORIZONTAL)
        self.m_staticTextx7 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Details of Change:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticTextx7.Wrap(0)
        bSizerx17.Add(self.m_staticTextx7, 0, wx.ALL, 5)

        self.m_textCtrl1x = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 85),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )
        bSizerx17.Add(self.m_textCtrl1x, 5, wx.ALL, 5)
        self.m_textCtrl1x.SetEditable(False)
        self.m_staticTextx7x = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Reason for Change:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticTextx7x.Wrap(0)
        bSizerx17.Add(self.m_staticTextx7x, 0, wx.ALL, 5)

        self.m_textCtrl1xx = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 85),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )
        self.m_textCtrl1xx.SetEditable(False)
        bSizerx17.Add(self.m_textCtrl1xx, 5, wx.ALL, 5)
        bSizer3.Add(bSizerx17, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer5.AddSpacer(50)

        self.sign_m_button3 = wx.Button(
            self,
            wx.ID_ANY,
            u'Confirm you have digitally signed and saved the document',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer5.Add(self.sign_m_button3, 0, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer5.AddSpacer(50)

        bSizer3.Add(bSizer5, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL,
                    5)

        bSizer6 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer3.Add(bSizer6, 1, wx.EXPAND, 5)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.sign_m_staticText21 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Reject PEOI if not OK',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.sign_m_staticText21.Wrap(-1)
        self.sign_m_staticText21.SetFont(wx.Font(
            10,
            75,
            90,
            92,
            False,
            'Consolas',
        ))

        bSizer101.Add(self.sign_m_staticText21, 0, wx.ALL
                      | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer3.Add(bSizer101, 0, wx.EXPAND, 5)

        bSizer17 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText7 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Reason for Rejection:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText7.Wrap(0)
        bSizer17.Add(self.m_staticText7, 0, wx.ALL, 5)

        self.m_textCtrl1 = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 85),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )
        bSizer17.Add(self.m_textCtrl1, 5, wx.ALL, 5)

        bSizer3.Add(bSizer17, 0, wx.EXPAND, 5)

        bSizer171 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText71 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Rejected by:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText71.Wrap(0)
        bSizer171.Add(self.m_staticText71, 0, wx.ALL, 5)

        self.m_textCtrl11 = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(-1, 30),
            wx.TE_CHARWRAP | wx.TE_MULTILINE,
        )
        bSizer171.Add(self.m_textCtrl11, 2, wx.ALL, 5)

        self.m_button4 = wx.Button(
            self,
            wx.ID_ANY,
            u'Reject',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer171.Add(self.m_button4, 1, wx.ALL, 5)

        bSizer3.Add(bSizer171, 0, wx.EXPAND, 5)

        if wordy.userlevel == 'QA' or wordy.userlevel == 'PD':
            self.fillfields.Disable()
            bSizer4.Hide(self.fillfields)

        self.m_comboBox2.Bind(wx.EVT_COMBOBOX,
                              self.m_comboBox2OnCombobox)
        self.sign_m_button2.Bind(
            wx.EVT_BUTTON, lambda event: self.approvePDF(event, 'CU'))
        self.sign_m_button2AR.Bind(
            wx.EVT_BUTTON, lambda event: self.approvePDF(event, 'AR'))
        self.sign_m_button3.Bind(wx.EVT_BUTTON, self.accept_PEOI)
        self.fillfields.Bind(wx.EVT_BUTTON, self.fill_fields_PDF)
        self.m_button4.Bind(wx.EVT_BUTTON, self.reject_PEOI)
        self.fname = self.m_comboBox2.GetValue()

        self.SetSizer(bSizer3)
        self.Layout()
        self.Show()

    def filterIT(self, lizt):
        lizzt = []
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        for PEOIx in lizt:
            (REV, PEOI) = getRnP(PEOIx)
            found_you = findme(self.roctop, PEOI)
            if found_you >= 0:
                found2 = moreFinding(self.roctop[found_you], REV)
                if found2 > 0:
                    signer = safestr(wordy.origin) + '_sign'
                    if self.roctop[found_you][found2]['revs'][signer] \
                            == 0:
                        lizzt.append(PEOIx)
                    elif self.roctop[found_you][found2]['revs'][signer] \
                            == 1:
                        pass
                    elif self.roctop[found_you][found2]['revs'][signer] \
                            == 2:
                        lizzt.append(PEOIx)
                    if self.roctop[found_you][found2]['revs']['PE_sign'
                                                              ] == 1:
                        if self.roctop[found_you][found2]['revs'
                                                          ]['PD_sign'] == 1:
                            if self.roctop[found_you][found2]['revs'
                                                              ]['QA_sign'] == 1:
                                rej_file = \
                                    os.path.abspath(wordy.forappralpath
                                                    + '\\' + PEOI + '_' + REV + '\\'
                                                    + PEOI + '_' + REV + '.txt')
                                if os.path.isfile(rej_file):
                                    if fileaccesscheck(rej_file):
                                        os.remove(rej_file)
        return lizzt

    def load_form(self, filename):
        """Load pdf form contents into a nested list of name/value tuples"""
        with open(filename, 'rb') as file:
            parser = PDFParser(file)
            doc = PDFDocument(parser)
            parser.set_document(doc)
            return [self.load_fields(resolve1(f)) for f in
                    resolve1(doc.catalog['AcroForm'])['Fields']]

    def load_fields(self, field):
        return (field.get('T').decode('utf-8'), resolve1(field.get('V')))

    def get_sig_fields(self):
        fpath = os.path.join(wordy.forappralpath + "\\" +
                             wordy.PEOIx[:-4] + "\\" + wordy.PEOIx)
        if os.path.isfile(fpath):
            f = PdfFileReader(fpath)
            ff = f.getFields()
            sigs = []
            flds = []
            try:
                for x in ff:
                    flds.append(str(x))
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                pass
            if len(flds) > 0:
                form = self.load_form(fpath)
                for x in form:
                    if isinstance(x, tuple):
                        for z in x:
                            if isinstance(z, dict):
                                try:
                                    sigs.append(z["Name"])
                                except Exception as e:
                                    print('oops ' + str((inspect.stack()[0][2])))
                                    print (e.message, e.args)
                                    sigs.append("unsigned")
                return [flds, sigs]
            return False
        else:
            return False

    def signed_off_PDF(self, r_or_a):
        signer = safestr(wordy.origin) + '_sign'
        signers_sig = safestr(wordy.origin) + '_sign_name'
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        if self.m_textCtrl1x.GetLabel() != '':
            found_you = findme(self.roctop, self.PEOI)
            if found_you >= 0:
                found2 = moreFinding(self.roctop[found_you], self.REV)

                if found2 > 0:
                    result = self.roctop[found_you][found2].get(
                        'revs', {}).get(signers_sig)
                    if r_or_a == 1:
                        pdfcount = self.get_sig_fields()
                        if pdfcount:
                            PDFval = int(len(pdfcount[1]))
                            sigstring = ", "
                            sigsfound = sigstring.join(pdfcount[1])
                            sigs_in_PDF = pdfcount[1]
                        else:
                            PDFval = 0
                            sigsfound = "None"
                            sigs_in_PDF = []
                        # compare these two and put this in signers sig
                        sigs_in_DB = [(self.roctop[found_you][found2].get('revs', {}).get('QA_sign_name')),
                                      (self.roctop[found_you][found2].get(
                                          'revs', {}).get('PE_sign_name')),
                                      (self.roctop[found_you][found2].get('revs', {}).get('PD_sign_name'))]
                        while None in sigs_in_DB:
                            sigs_in_DB.remove(None)
                        while None in sigs_in_PDF:
                            sigs_in_DB.remove(None)
                        if len(sigs_in_PDF) - len(sigs_in_DB) == 1:
                            wesult = list(set(sigs_in_PDF) - set(sigs_in_DB))
                            if len(wesult) == 1:
                                result = wesult[0]

                        DBval = int(self.get_count_of_DB_sigs(
                            self.roctop[found_you][found2]['revs']))
                        if (int(PDFval) - int(DBval)) >= 1:
                            a_or_r = 'you have approved ' + self.PEOI + '-' + self.REV
                        else:
                            badmessage = "I don't think you signed the PDF. Only these signatures were found: %s" % (
                                sigsfound)
                            dlg = wx.MessageDialog(
                                None, badmessage, "oops", wx.OK)
                            dlg.ShowModal()
                            dlg.Destroy()
                            return
                    else:
                        a_or_r = 'you have rejected ' + self.PEOI + '-' + self.REV
                        self.roctop[found_you][found2]['revs']['QA_sign'] = 0
                        self.roctop[found_you][found2]['revs']['PE_sign'] = 0
                        self.roctop[found_you][found2]['revs']['PD_sign'] = 0

                    self.roctop[found_you][found2]['revs'][signer] = r_or_a
                    self.roctop[found_you][found2]['revs'][signers_sig] = result

                    self.dodump()
                    dlg = wx.MessageDialog(None,
                                           'Record of Change Database updated',
                                           a_or_r, wx.OK)

                    dlg.ShowModal()
                    dlg.Destroy()
        else:

            dlg = wx.MessageDialog(None,
                                   'Cannot approve / reject without the Record of Change details for this revision level being completed (DOC and RFC are empty)', 'error', wx.OK)
            dlg.ShowModal()
            dlg.Destroy()

    def dodump(self):
        with open(self.theDB, 'wb') as f:
            PKL_dump(self.roctop, f, indent = 2)
        print(datetime.datetime.now())
        return

    def get_count_of_DB_sigs(self, dict_val):
        dbc = 0
        listihaventmadeyet = ['PE_sign', 'PD_sign', 'QA_sign']
        for n in range(0, 3):
            if int(dict_val[listihaventmadeyet[n]]) == 1:
                dbc += 1
        return dbc

    def load_data(self, PEOIx):
        wordy.PEOIx = PEOIx
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return
        found_you = -1
        found2 = -1
        (REV, PEOI) = getRnP(PEOIx)
        found_you = findme(self.roctop, PEOI)
        if found_you >= 0:
            found2 = moreFinding(self.roctop[found_you], REV)
            if found2 > 0:
                try:
                    DOC = safestr(self.roctop[found_you][found2]['revs'
                                                                 ]['doc'])
                    RFC = safestr(self.roctop[found_you][found2]['revs'
                                                                 ]['rfc'])
                    self.m_textCtrl1x.SetLabel(DOC)
                    self.m_textCtrl1xx.SetLabel(RFC)
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    pass
                signer = safestr(wordy.origin) + '_sign'
                if self.roctop[found_you][found2]['revs'][signer] == 0 or self.roctop[found_you][found2]['revs'][signer] == '':
                    self.sign_m_staticText211.SetLabel(PEOI + '-' + REV
                                                       + ' needs approving / rejecting by '
                                                       + wordy.origin)
                elif self.roctop[found_you][found2]['revs'][signer] == 1:
                    self.sign_m_staticText211.SetLabel(PEOI + '-' + REV
                                                       + ' has been marked as approved by '
                                                       + wordy.origin)
                elif self.roctop[found_you][found2]['revs'][signer] == 2:
                    self.sign_m_staticText211.SetLabel(PEOI + '-' + REV
                                                       + ' has been marked as rejected by '
                                                       + wordy.origin)
            self.REV = REV
            self.PEOI = PEOI
        else:
            self.m_textCtrl1x.SetLabel('')
            self.m_textCtrl1xx.SetLabel('')

    def accept_PEOI(self, event):
        if self.fname != '':
            self.signed_off_PDF(1)
            self.load_data(self.fname)
            wordy.suppress = 0
            self.Close()
            x = ApproveFrame('Approver (re-opened)')
        return

    def reject_PEOI(self, event):
        if self.m_textCtrl1x.GetLabel() != '':
            if self.m_textCtrl1.GetValue() != '':
                if self.m_textCtrl11.GetValue() != '':
                    self.signed_off_PDF(2)
                    if self.fname != '':
                        fpath = self.fname[: len(self.fname) - 4]
                        out_file = os.path.abspath(wordy.forappralpath
                                                   + '\\' + fpath + '\\' + fpath + '.txt')
                        with open(out_file, 'a+') as f:
                            f.write(self.PEOI + '-' + self.REV
                                    + ' has been rejected by '
                                    + wordy.origin)
                            f.write('\n')
                            f.write(self.m_textCtrl1.GetValue())
                            f.write('\n')
                            f.write(self.m_textCtrl11.GetValue())
                            f.write('\n')
                    self.load_data(self.fname)
                    return
                else:
                    dlg = wx.MessageDialog(None,
                                           "please fill in 'rejected by' field",
                                           'error', wx.OK)
            else:
                dlg = wx.MessageDialog(None,
                                       "please fill in 'reason for rejection' field",
                                       'error', wx.OK)
        else:
            dlg = wx.MessageDialog(None,
                                   'Cannot reject without the Record of Change details for this revision level being completed (DOC and RFC are empty)', 'error', wx.OK)
        dlg.ShowModal()
        dlg.Destroy()
        wordy.suppress = 0
        self.Close()
        x = ApproveFrame('Approver (re-opened)')
        return

    def getPDFexepath(self, tofind, alt):
        try:
            os.system('mode con: cols=15 lines=9999')
            k = r'Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\*"  | Where-Object {$q + $_."(default)" -ne $null} | Select-Object @{ expression={$_.PSChildName}; label="Program"} ,@{ expression={$q + $_."(default)" +$q}; label="CommandLine"} | Export-Csv -Path $env:temp\programs.csv -Encoding ascii -NoTypeInformation'
            p = subprocess.run(["powershell", "-Command", k])
            exepath = None
            temp_dir = gettempdir()
            execeslist = os.path.abspath(temp_dir  + '\\' + 'programs.csv')
            with open(execeslist) as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                        if tofind in row:
                            exepath = row[1]
            if exepath == None:
                f = r'Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" |  Export-Csv -Path $env:temp\programs2.csv -Encoding ascii -NoTypeInformation'
                p = subprocess.run(["powershell", "-Command", f])
                execeslist2 = os.path.abspath(temp_dir  + '\\' + 'programs2.csv')
                with open(execeslist2) as csvfile:
                    reader = csv.reader((line.replace('\0', '') for line in csvfile), delimiter=",")
                    for row in reader:
                        for cell in row:
                            if tofind in cell:
                                exepath = cell
            if exepath is not None:
                print ('found it - ' + exepath)
                return exepath
            else:
                print ("can't find it, taking a guess")
                sleep(10)
                return alt
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            print ("really can't find it! Taking a guess")
            sleep(10)
            return alt

    def fill_fields_PDF(self, event):
        if self.fname != '':
            fpath = self.fname[: len(self.fname) - 4]
            out_file = os.path.abspath(wordy.forappralpath + '\\'
                                       + fpath + '\\' + self.fname)
            try:
                if os.path.isfile(out_file):
                    Fcall("PDF_fields.exe " + out_file, shell=True)
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                pass
        return

    def approvePDF(self, event, opn):
        if self.fname != '':
            fpath = self.fname[: len(self.fname) - 4]
            out_file = os.path.abspath(wordy.forappralpath + '\\'
                                       + fpath + '\\' + self.fname)
        if opn == 'CU':
            c_or_a = 'cpdf_install'
        elif opn == 'AR':
            c_or_a = 'apdf_install'
        config = configparser.ConfigParser()
        config.read('settings.ini')
        try:
            acrobat_path = config['user_profile'][c_or_a]
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            acrobat_path = 'U'
        if acrobat_path == 'U':
            msg = wx.BusyInfo(
                'Please wait, looking for install directory... this will only happen once, next time will be quicker')
            if opn == 'CU':
                acro_details = self.getPDFexepath('CutePDF.exe', r"C:\Program Files (x86)\Acro Software\CutePDF Pro\CutePDF.exe")
            elif opn == 'AR':
                acro_details = self.getPDFexepath('Acrobat.exe', r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe")
            try:
                config.set('user_profile', c_or_a, acro_details)
                with open('settings.ini', 'w') as configfile:
                    config.write(configfile)
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                print("config oops")
                pass
        elif acrobat_path == 'N':
            acro_details = 'not a directory'
        else:
            acro_details = config['user_profile'][c_or_a]

        try:
            config.read('settings.ini')
            os.system('mode con: cols=%s lines=%s'
                        % (config['consolesize']['cols'],
                            config['consolesize']['lines']))
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            os.system('mode con: cols=15 lines=15')

        if os.path.isfile(acro_details):
            SP_Popen([acro_details, out_file])
        else:
            try:
                config.set('user_profile', c_or_a, acro_details)
                with open('settings.ini', 'w') as configfile:
                    config.write(configfile)
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                print("config oops")
                pass
            print("cannot find the program, trying to open with Windows default instead")
            if os.path.isfile(out_file):
                SP_Popen(out_file, shell=True)
        return

    def m_comboBox2OnCombobox(self, event):
        fname = self.m_comboBox2.GetValue()
        if fname != '':
            if fname != self.fname:
                msg1 = 'Review %s and sign off if OK' % safestr(fname)
                msg2 = 'Reject %s if not OK' % safestr(fname)
                self.sign_m_staticText211.SetLabel(msg1)
                self.sign_m_staticText21.SetLabel(msg2)
                self.m_textCtrl1.SetLabel('')
                self.load_data(fname)
            self.fname = fname
        return

    def __del__(self):
        wordy.app_frame_number = 0


class OtherFrame(wx.Frame):

    """
    Class used for creating frames other than the main one
    """

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title=title,
            pos=wx.DefaultPosition,
            size=wx.Size(1024, 570),
            style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU |
            wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN | wx.TAB_TRAVERSAL,
        )
        
        self.SetIcon(OICC_LOGO_2.GetIcon())
        self.panel = wx.Panel(self)

        self.theDB = wordy.DBpath
        

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        b64_img_str = \
            'eJxz8mVjZgADMyDWAGJ+KGZkkGCAgSN8EAwDH0gBEPXGxAFk9f8JAaqo//btWzYYABkE1UMUe4ABphY09XDFj8AAUwua+h07dkAUQ2QxtRD0L5oWYsIHWQuR4QnRAnQtneOLpPRDUvoEACz7/r8='
        self.tickimage = wx.Image(self.create_bitstream_img(
            b64_img_str), wx.BITMAP_TYPE_ANY)

        b64_img_str = \
            'eJxz8mVjZgADMyDWAGJ+KGZkkGCAgSN8EAwDH0gBEPXGxAFk9f/////0/j9WABFHU//j+38vpf971qIrBooAxYGymOafP/LfQQxFC5ANFAGKY5oPUYCsBVkxLvVwLf1lKIrxqAeC2a3/Tdj/z2lF8Qgu9RBnrJqO7hes6pHdjOZ9TPVoHkTTgqYeGCnAcEZWDNcCFAfKYpoPjBSsACKOpp6k9ENS+gQAXaT6mg=='
        self.crossimage = wx.Image(
            self.create_bitstream_img(b64_img_str), wx.BITMAP_TYPE_ANY)

        b64_img_str = \
            'eJxz8mVjZgADMyDWAGJ+KGZkkGCAgSN8EAwDH0gBEPXGxAFk9f8JgVH1I0o9SemHpPQJAP3UJrU='
        self.awaitimage = wx.Image(
            self.create_bitstream_img(b64_img_str), wx.BITMAP_TYPE_ANY)

        self.imgbank = [self.awaitimage, self.awaitimage, self.awaitimage,
                        self.awaitimage, self.tickimage, self.crossimage]
        self.awaitr = []
        for n in range(0, 6):
            self.awaitr.append(wx.StaticBitmap(self.panel, id=wx.ID_ANY, bitmap=wx.Bitmap(
                self.imgbank[n]), pos=wx.DefaultPosition, size=(14, 14), style=0))
        self.m_bitmap2 = self.awaitr[0]
        self.m_bitmap1 = self.awaitr[1]
        self.m_bitmap0 = self.awaitr[2]
        self.m_bitmapAWAIT = self.awaitr[3]
        self.m_bitmapOK = self.awaitr[4]
        self.m_bitmapREJ = self.awaitr[5]

        self.roctop = []
        if wordy.reset_after_reimport == 1:
            wordy.reset_after_reimport = 0
            self.resetter()

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.ROC2OPIN = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Record of Change to Operator Instructions',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.ROC2OPIN.Wrap(-1)
        self.ROC2OPIN.SetFont(wx.Font(
            14,
            74,
            90,
            90,
            False,
            'Arial',
        ))

        bSizer2.Add(self.ROC2OPIN, 0, wx.ALL, 5)

        self.m_staticline1 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL)
        bSizer2.Add(self.m_staticline1, 0, wx.EXPAND | wx.ALL, 5)

        self.stat_PEOI_num = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Instruction Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_PEOI_num.Wrap(-1)
        bSizer2.Add(self.stat_PEOI_num, 0, wx.ALL, 5)

        self.m_textCtrl_PEOI = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        if wordy.userlevel != 'master':
            self.m_textCtrl_PEOI.SetEditable(False)
            self.m_textCtrl_PEOI.SetCursor(wx.Cursor(wx.CURSOR_HAND))
            self.m_textCtrl_PEOI.SetBackgroundColour((195, 195, 195))
        bSizer2.Add(self.m_textCtrl_PEOI, 1, wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_cust_pn = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Customer and part Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_cust_pn.Wrap(-1)
        bSizer3.Add(self.stat_cust_pn, 0, wx.ALL, 5)

        self.m_textCtrl_cust_PN = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        wx.CallAfter(self.m_textCtrl_cust_PN.SetInsertionPoint, 0)
        bSizer3.Add(self.m_textCtrl_cust_PN, 1, wx.ALL, 5)

        self.m_staticline2 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer3.Add(self.m_staticline2, 0, wx.EXPAND | wx.ALL, 5)

        self.stat_Pek_PN = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Pektron Part Number',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_Pek_PN.Wrap(-1)
        bSizer3.Add(self.stat_Pek_PN, 0, wx.ALL, 5)

        self.m_textCtrl_pekPN = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer3.Add(self.m_textCtrl_pekPN, 1, wx.ALL, 5)

        bSizer1.Add(bSizer3, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_prod_desc = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Product Description',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_prod_desc.Wrap(-1)
        bSizer4.Add(self.stat_prod_desc, 0, wx.ALL, 5)

        self.m_textCtrl_prod_desc = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer4.Add(self.m_textCtrl_prod_desc, 1, wx.ALL, 5)

        self.stat_stages = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Stage numbers',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_stages.Wrap(-1)
        bSizer4.Add(self.stat_stages, 0, wx.ALL, 5)

        self.m_textCtrl_stages = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer4.Add(self.m_textCtrl_stages, 0, wx.ALL, 5)

        bSizer1.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_process_desc = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Operation / Process',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_process_desc.Wrap(-1)
        bSizer5.Add(self.stat_process_desc, 0, wx.ALL, 5)

        self.m_textCtrl_proc = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        
        self.isobs = wx.CheckBox(
            self.panel,
            wx.ID_ANY,
            u"Obsolete? (not the\nRevision, the entire PEOI)",
            wx.DefaultPosition,
            wx.DefaultSize,
            0,                       
        )
        
        bSizer5.Add(self.m_textCtrl_proc, 1, wx.ALL, 5)
        bSizer5.Add(self.isobs, 0, wx.ALL, 5)

        bSizer1.Add(bSizer5, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline3 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer6.Add(self.m_staticline3, 1, wx.ALL | wx.EXPAND, 5)

        bSizer1.Add(bSizer6, 0, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        self.stat_ISSUE = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'ISSUE',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_ISSUE.Wrap(-1)
        bSizer8.Add(self.stat_ISSUE, 0, wx.ALL
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_ISS = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(60, 50),
            0,
        )
        if wordy.userlevel != 'master':
            self.m_textCtrl_ISS.SetEditable(False)
            self.m_textCtrl_ISS.SetBackgroundColour((195, 195, 195))
        bSizer8.Add(self.m_textCtrl_ISS, 0, wx.ALL, 5)

        bSizer7.Add(bSizer8, 0, 0, 5)

        bSizer81 = wx.BoxSizer(wx.VERTICAL)

        self.stat_DOC = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Detail of Change',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_DOC.Wrap(-1)
        bSizer81.Add(self.stat_DOC, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_DOC = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(200, 120),
            wx.TE_MULTILINE,
        )
        bSizer81.Add(self.m_textCtrl_DOC, 0, wx.ALL, 5)

        bSizer7.Add(bSizer81, 0, 0, 5)

        bSizer811 = wx.BoxSizer(wx.VERTICAL)

        self.stat_RFC = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Reason for Change',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_RFC.Wrap(-1)
        bSizer811.Add(self.stat_RFC, 0, wx.ALL
                      | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_RFC = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(200, 120),
            wx.TE_MULTILINE,
        )
        bSizer811.Add(self.m_textCtrl_RFC, 0, wx.ALL, 5)

        bSizer7.Add(bSizer811, 0, 0, 5)

        bSizer19 = wx.BoxSizer(wx.VERTICAL)

        bSizer35 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer8111 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_pages = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'pages',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_pages.Wrap(-1)
        bSizer8111.Add(self.stat_pages, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL
                       | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl_pages = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(50, 25),
            0,
        )
        bSizer8111.Add(self.m_textCtrl_pages, 0, wx.ALL
                       | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer35.Add(bSizer8111, 0, 0, 5)

        bSizer81111 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_impl = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'implemented by',
            wx.DefaultPosition,
            wx.Size(75, 40),
            0,
        )
        self.stat_impl.Wrap(-1)
        bSizer81111.Add(self.stat_impl, 0, wx.ALL
                        | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_impl = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(200, 25),
            0,
        )
        bSizer81111.Add(self.m_textCtrl_impl, 0, wx.ALL
                        | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer35.Add(bSizer81111, 0, 0, 5)

        bSizer19.Add(bSizer35, 1, wx.EXPAND, 5)

        bSizer81113 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer81112 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_copies = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'copies',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_copies.Wrap(-1)
        bSizer81112.Add(self.stat_copies, 0, wx.ALL
                        | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_copies = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(50, 25),
            0,
        )
        bSizer81112.Add(self.m_textCtrl_copies, 0, wx.ALL, 5)

        bSizer81113.Add(bSizer81112, 0, 0, 5)

        bSizer811111 = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_rcvd = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'received by',
            wx.DefaultPosition,
            wx.Size(75, -1),
            0,
        )
        self.stat_rcvd.Wrap(-1)
        bSizer811111.Add(self.stat_rcvd, 0, wx.ALL
                         | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_rcvd = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(200, 25),
            0,
        )
        bSizer811111.Add(self.m_textCtrl_rcvd, 0, wx.ALL, 5)

        bSizer81113.Add(bSizer811111, 0, 0, 5)

        bSizer19.Add(bSizer81113, 0, 0, 5)

        bSizer56 = wx.BoxSizer(wx.HORIZONTAL)
        bSizer56A = wx.BoxSizer(wx.HORIZONTAL)
        bSizer56B = wx.BoxSizer(wx.HORIZONTAL)
        bSizer56C = wx.BoxSizer(wx.HORIZONTAL)
        bSizer56D = wx.BoxSizer(wx.HORIZONTAL)

        self.m_reset = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'RST',
            wx.DefaultPosition,
            wx.Size(30, 25),
            0,
        )

        self.PE_sig_name = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'PE Sign name:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.QA_sig_name = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'QA Sign name:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.PD_sig_name = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'PD Sign name:',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.m_checkBoxPE = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Prod Eng signed',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        bSizer56.Add(self.m_reset, 1, wx.ALL | wx.ALIGN_LEFT
                     | wx.EXPAND, 5)
        bSizer56A.Add(self.m_checkBoxPE, 1, wx.ALL | wx.ALIGN_RIGHT
                      | wx.EXPAND, 5)
        bSizer56A.Add(self.m_bitmap0, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)

        self.m_checkBoxQA = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'QA signed',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        bSizer56B.Add(self.m_checkBoxQA, 1, wx.ALL | wx.ALIGN_RIGHT
                      | wx.EXPAND, 5)
        bSizer56B.Add(self.m_bitmap1, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)

        self.m_checkBoxPN = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Prod Signed',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        bSizer56C.Add(self.m_checkBoxPN, 1, wx.ALL | wx.ALIGN_RIGHT
                      | wx.EXPAND, 5)
        bSizer56C.Add(self.m_bitmap2, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)

        bSizer56D.Add(self.PE_sig_name, 1, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        bSizer56D.Add(self.QA_sig_name, 1, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        bSizer56D.Add(self.PD_sig_name, 1, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)

        bSizer56.Add(bSizer56A, 0, wx.ALL | wx.EXPAND, 5)
        bSizer56.Add(bSizer56B, 0, wx.ALL | wx.EXPAND, 5)
        bSizer56.Add(bSizer56C, 0, wx.ALL | wx.EXPAND, 5)

        bSizer19.Add(bSizer56, 0, wx.EXPAND, 5)
        bSizer19.Add(bSizer56D, 0, wx.ALL | wx.EXPAND, 5)

        bSizer7.Add(bSizer19, 1, 0, 5)

        bSizer82 = wx.BoxSizer(wx.VERTICAL)

        bSizer82XX = wx.BoxSizer(wx.HORIZONTAL)

        self.stat_dateRC = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Date Received',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.stat_dateRC.Wrap(-1)
        bSizer82.Add(self.stat_dateRC, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl_dateRC = wx.TextCtrl(
            self.panel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.Size(75, 25),
            0,
        )
        bSizer82.Add(self.m_textCtrl_dateRC, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_button2 = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'submit',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        self.m_button_previous = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'<',
            wx.DefaultPosition,
            wx.Size(20, 25),
            0,
        )

        self.m_button_next = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'>',
            wx.DefaultPosition,
            wx.Size(20, 25),
            0,
        )
        self.m_button_Sprevious = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'<S',
            wx.DefaultPosition,
            wx.Size(20, 25),
            0,
        )

        self.m_button_Snext = wx.Button(
            self.panel,
            wx.ID_ANY,
            u'S>',
            wx.DefaultPosition,
            wx.Size(20, 25),
            0,
        )
        bSizer82.Add(self.m_button2, 0, wx.ALL
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer82XX.Add(self.m_button_previous, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer82XX.Add(self.m_button_Sprevious, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer82XX.Add(self.m_button_Snext, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer82XX.Add(self.m_button_next, 0, wx.ALL
                       | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer82.Add(bSizer82XX, 1, wx.EXPAND, 5)

        bSizer7.Add(bSizer82, 1, wx.EXPAND, 5)

        bSizer1.Add(bSizer7, 0, wx.EXPAND, 5)

        bSizer61 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline31 = wx.StaticLine(self.panel, wx.ID_ANY,
                                            wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer61.Add(self.m_staticline31, 0, wx.ALL | wx.EXPAND, 5)

        bSizer52 = wx.BoxSizer(wx.VERTICAL)
        bSizer52a = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText31 = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Previous issue levels (if any - max last five)                  ',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText31.Wrap(-1)
        self.m_staticText31A = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'        Awaiting sign-off = ',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        # self.m_staticText31A.Wrap( -1 )

        self.m_staticText31B = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'        Signed off = ',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        # self.m_staticText31B.Wrap( -1 )

        self.m_staticText31C = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'        Rejected = ',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )

        # self.m_staticText31C.Wrap( -1 )

        bSizer52a.Add(self.m_staticText31, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        self.m_staticline31a = wx.StaticLine(self.panel, wx.ID_ANY,
                                             wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL)
        bSizer52a.Add(self.m_staticline31a, 0, wx.EXPAND | wx.ALL, 5)
        bSizer52a.Add(self.m_staticText31A, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        bSizer52a.Add(self.m_bitmapAWAIT, 0, wx.ALL
                      | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)
        bSizer52a.Add(self.m_staticText31B, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        bSizer52a.Add(self.m_bitmapOK, 0, wx.ALL
                      | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)
        bSizer52a.Add(self.m_staticText31C, 0, wx.ALL | wx.ALIGN_LEFT
                      | wx.EXPAND, 5)
        bSizer52a.Add(self.m_bitmapREJ, 0, wx.ALL
                      | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)
        bSizer52.Add(bSizer52a, 0, wx.ALL | wx.ALIGN_RIGHT | wx.EXPAND,
                     5)
        self.roc_prev = WX_Grid(self.panel, wx.ID_ANY,
                                wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid

        self.roc_prev.CreateGrid(7, 8)
        self.roc_prev.EnableEditing(False)
        self.roc_prev.EnableGridLines(True)
        self.roc_prev.EnableDragGridSize(False)
        self.roc_prev.SetMargins(0, 0)

        # Columns

        self.roc_prev.SetColSize(0, 60)
        self.roc_prev.SetColSize(1, 274)
        self.roc_prev.SetColSize(2, 274)
        self.roc_prev.SetColSize(3, 60)
        self.roc_prev.SetColSize(4, 60)
        self.roc_prev.SetColSize(5, 90)
        self.roc_prev.SetColSize(6, 90)
        self.roc_prev.SetColSize(7, 90)
        self.roc_prev.EnableDragColMove(False)
        self.roc_prev.EnableDragColSize(True)
        self.roc_prev.SetColLabelSize(30)
        self.roc_prev.SetColLabelValue(0, u'ISSUE')
        self.roc_prev.SetColLabelValue(1, u'DETAILS')
        self.roc_prev.SetColLabelValue(2, u'REASON')
        self.roc_prev.SetColLabelValue(3, u'PGS')
        self.roc_prev.SetColLabelValue(4, u'COPIES')
        self.roc_prev.SetColLabelValue(5, u'IMPLMNT')
        self.roc_prev.SetColLabelValue(6, u'RCVD')
        self.roc_prev.SetColLabelValue(7, u'DATE')
        self.roc_prev.SetColLabelAlignment(wx.ALIGN_CENTRE,
                                           wx.ALIGN_CENTRE)

        # Rows

        self.roc_prev.EnableDragRowSize(False)
        self.roc_prev.SetRowLabelSize(0)
        self.roc_prev.SetRowLabelAlignment(wx.ALIGN_CENTRE,
                                           wx.ALIGN_CENTRE)

        # Cell Defaults

        self.roc_prev.SetDefaultCellAlignment(wx.ALIGN_LEFT,
                                              wx.ALIGN_TOP)
        bSizer52.Add(self.roc_prev, 1, wx.ALL | wx.EXPAND, 5)

        bSizer61.Add(bSizer52, 1, wx.EXPAND
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer61, 1, wx.EXPAND
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)
        self.m_button2.Bind(wx.EVT_BUTTON, self.submitter)
        self.isobs.Bind(wx.EVT_CHECKBOX,self.obset)
        self.m_reset.Bind(wx.EVT_BUTTON, self.reset_approvals)
        self.m_button_previous.Bind(
            wx.EVT_BUTTON, lambda event: self.load_next(event, -1))
        self.m_button_next.Bind(
            wx.EVT_BUTTON, lambda event: self.load_next(event, 1))
        self.m_button_Sprevious.Bind(
            wx.EVT_BUTTON, lambda event: self.load_NSAVE_next(event, -1))
        self.m_button_Snext.Bind(
            wx.EVT_BUTTON, lambda event: self.load_NSAVE_next(event, 1))
        if wordy.userlevel == 'master':
            self.popupmenu2 = wx.Menu()
            for text2 in ['aw_app', 'approved', 'rejected']:
                item2 = self.popupmenu2.Append(-1, text2)
                self.Bind(wx.EVT_MENU, self.OnPopupItemSelected2, item2)
            self.PE_sign = 0
            self.PD_sign = 0
            self.QA_sign = 0
            self.m_bitmap0.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
                event, self.m_bitmap0, self.popupmenu2, 'PE'))
            self.m_bitmap1.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
                event, self.m_bitmap1, self.popupmenu2, 'QA'))
            self.m_bitmap2.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
                event, self.m_bitmap2, self.popupmenu2, 'PD'))

        self.popupmenu = wx.Menu()
        menulist = ['view the PDF', 'paste', '______', 'today', 'me', 'First Issue', '______', 'set me']
        if wordy.headerstuff:
            if wordy.headerstuff[0]==wordy.extractedID[0]:
                menulist=menulist+['______','guess all', '______', 'Prod desc guess', 'Cust PN guess', 'Stages guess', 'Stage desc guess', 'Pek PN guess']
        
        for text in menulist:
            item = self.popupmenu.Append(-1, text)
            self.Bind(wx.EVT_MENU, self.OnPopupItemSelected, item)
            
        self.m_textCtrl_PEOI.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_PEOI, self.popupmenu, None))
        self.m_textCtrl_ISS.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_ISS, self.popupmenu, None))
        self.m_textCtrl_cust_PN.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_cust_PN, self.popupmenu, None))
        self.m_textCtrl_proc.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_proc, self.popupmenu, None))
        self.m_textCtrl_pekPN.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_pekPN, self.popupmenu, None))
        self.m_textCtrl_prod_desc.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_prod_desc, self.popupmenu, None))
        self.m_textCtrl_stages.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_stages, self.popupmenu, None))
        self.m_textCtrl_DOC.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_DOC, self.popupmenu, None))
        self.m_textCtrl_RFC.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_RFC, self.popupmenu, None))
        self.m_textCtrl_pages.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_pages, self.popupmenu, None))
        self.m_textCtrl_impl.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_impl, self.popupmenu, None))
        self.m_textCtrl_copies.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_copies, self.popupmenu, None))
        self.m_textCtrl_rcvd.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_rcvd, self.popupmenu, None))
        self.m_textCtrl_dateRC.Bind(wx.EVT_CONTEXT_MENU, lambda event: self.on1Focus(
            event, self.m_textCtrl_dateRC, self.popupmenu, None))
        self.panel.SetSizer(bSizer1)
        self.panel.Layout()

        self.panel.Centre(wx.BOTH)

        self.m_textCtrl_PEOI.SetLabel(wordy.extractedID[0])
        self.m_textCtrl_ISS.SetLabel(wordy.extractedID[1])

        self.load_data()
        if wordy.do_save == 1:
            self.submitter(None)
            wordy.do_save = 0
            self.load_data()
        self.done_this = 0
        self.panel.Show()
        self.Show()
        
        
    def obset(self, event):                
        if self.isobs.GetValue():
            tlist3 = []
            found_you = findme(self.roctop, wordy.extractedID[0])
            found2 = foolproof_finder(self.roctop[found_you], 1)
            if found2 > 0:
                Fu = 1
                while Fu <= found2:
                    tlist3.append(wordy.extractedID[0] + '_' + self.roctop[found_you][Fu]['revs']['rev'] + '.pdf')
                    Fu += 1
            #print(tlist3[0])
            locationz=[wordy.approvedpath, wordy.forappralpath]
            for docz in tlist3:
                for locz in locationz:
                    try:      
                        #pdf_name = wordy.extractedID[0] + '_' + wordy.extractedID[1] + '.pdf'
                        (ignore, current_doc) = search_file(locz, docz)
                        if os.path.isfile(current_doc):
                            apply_archive = wx.MessageDialog(None,
                                                ('PDF ' + current_doc + ' found in ' + ignore + '. This action makes the entire instruction obsolete.') + '\n' +
                                                ('If you click YES, the PDF signatures will become invalid, and it will.') + '\n' +
                                                ('be moved to the archive folder with the obsolete stamp applied.'),
                                                #('The PEOI number will be marked as obsolete in the DB.'), + '\n' +
                                                #('It will be marked as obsolete in OPindex and Printable ROCs.'), + '\n' +                                                                                      
                                                ('Cancel = NO, Proceed = YES'),
                                                wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION).ShowModal()
                        if apply_archive != wx.ID_YES:
                            pass
                        elif apply_archive == wx.ID_YES:
                            wordy.suppress = 1                    
                            wordy.m_filePicker1 = current_doc
                            wordy(None, title='duplicrap').m_filePicker1OnFileChanged(
                                wordy)
                            wordy.m_filePicker1 = None
                            wordy.suppress = 0
                            try:
                                wordy.do_save = 1
                                self.submitter(None)
                                wordy.do_save = 0
                            except Exception as e:
                                print('oops ' + str((inspect.stack()[0][2])))
                                print (e.message, e.args)
                                pass
                    except TypeError:
                        loc = locz.rsplit('\\', 1)[-1] or locz
                        print ("PDF " + docz + " not found in " + loc + " (that's OK, just checking)")
        self.onChecked(None)
        
        
    def no_edits(self):
        all_the_fields = [self.m_textCtrl_copies,
            self.m_textCtrl_cust_PN,
            self.m_textCtrl_rcvd,
            self.m_textCtrl_DOC,
            self.m_textCtrl_dateRC,
            self.m_textCtrl_RFC,
            self.m_textCtrl_impl,
            self.m_textCtrl_pages,
            self.m_textCtrl_proc,
            self.m_textCtrl_prod_desc,
            self.m_textCtrl_pekPN,
            self.m_textCtrl_stages]
        
        self.IDid = str(
            wordy.extractedID[0] + "_" + wordy.extractedID[1] + ".pdf")
        if not self.IDid in wordy.picklist:
            if not wordy.userlevel=="master":
                for field in all_the_fields:
                    field.SetEditable(False)
                    field.SetBackgroundColour((195, 195, 195))
        
    def onChecked(self, event):
        all_the_fields = [self.m_textCtrl_copies,
            self.m_textCtrl_cust_PN,
            self.m_textCtrl_rcvd,
            self.m_textCtrl_DOC,
            self.m_textCtrl_dateRC,
            self.m_textCtrl_RFC,
            self.m_textCtrl_impl,
            self.m_textCtrl_pages,
            self.m_textCtrl_proc,
            self.m_textCtrl_prod_desc,
            self.m_textCtrl_pekPN,
            self.m_textCtrl_stages]
        
        #cb = event.GetEventObject()
        cb = self.isobs
        cbn = not(cb.GetValue())
        for field in all_the_fields:
            field.SetEditable(cbn)
        
        if wordy.userlevel=="master":
            self.m_textCtrl_PEOI.SetEditable(cbn)
            self.m_textCtrl_ISS.SetEditable(cbn)

    def resetter(self):
        found_you = -1
        found2 = -1
        if not os.path.isfile(self.theDB):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return
        with open(self.theDB, 'rb') as f:
            roc2 = PKL_load(f)
        found_you = findme(roc2, wordy.extractedID[0])
        if found_you >= 0:
            if len(roc2[found_you]) > 0:
                findme2 = findingNemo(roc2[found_you],
                                      wordy.extractedID[1])
                found2 = findme2[0]
        self.roctop = roc2
        reset_stuff = {"PD_sign": 0, "QA_sign": 0, "PE_sign": 0,
                       "QA_sign_name": None, "PE_sign_name": None, "PD_sign_name": None}
        self.roctop[found_you][found2]['revs'].update(reset_stuff)
        self.saverfunc()
        self.Close()
        # now reset whatever needs resetting
        # then do the save
        return

    def create_bitstream_img(self, b64_img_str):
        image_data = decompress(b64decode(b64_img_str))
        stream = BytesIO(bytearray(image_data))
        return stream

    def load_NSAVE_next(self, event, up_or_down):
        try:
            wordy.do_save = 1
            self.submitter(None)
            wordy.do_save = 0
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        self.load_next(None, up_or_down)
        self.onChecked(None)
        return

    def load_next(self, event, up_or_down):
        try:
            self.IDid = str(
                wordy.extractedID[0] + "_" + wordy.extractedID[1] + ".pdf")
            wordy.posinlist = [i for i, s in enumerate(
                wordy.picklist) if self.IDid in s][0]
            if up_or_down == 1:
                if wordy.posinlist + 1 <= (len(wordy.picklist) - 1):
                    self.inf = wordy.picklist[wordy.posinlist + 1]
                else:
                    self.inf = wordy.picklist[0]
            elif up_or_down == -1:
                if wordy.posinlist - 1 >= 0:
                    self.inf = wordy.picklist[wordy.posinlist - 1]
                else:
                    self.inf = wordy.picklist[len(wordy.picklist) - 1]
            else:
                return
            if self.inf != '':
                REV, PEOI = getRnP(self.inf, True)
                lb = PEOI.find('-')
                if lb == 4:
                    wordy.extractedID = [safestr('PEOI-' + PEOI), REV,
                                         None]
            self.Close()
            title = 'Record of Change to Operator Instructions'
            frame = OtherFrame(title=title)
            wordy.frame_number = 1
        except IndexError:
            print("you can't skip forward / back on PEOI ROCs that are not 'in progress'")
            return
        found_you = findme(self.roctop, wordy.extractedID[0])
        if len(self.roctop[found_you]) > 0:
            findme2 = findingNemo(self.roctop[found_you], wordy.extractedID[1])
        found2 = findme2[0]
        self.dotickboxes(found_you, found2)
        self.m_textCtrl_rcvd.Disable()
        self.m_textCtrl_dateRC.Disable()
        if wordy.userlevel == "master":
            if self.QA_sign == 1:
                if self.PD_sign == 1:
                    if self.PE_sign == 1:
                        self.m_textCtrl_rcvd.Enable()
                        self.m_textCtrl_dateRC.Enable()
        self.onChecked(None)
        return

    def OnPopupItemSelected2(self, event):
        item = self.popupmenu2.FindItemById(event.GetId())
        text = item.GetText()
        if text == 'approved':
            self.res.SetBitmap(wx.Bitmap(self.tickimage))
            val_change = 1
        elif text == 'aw_app':
            self.res.SetBitmap(wx.Bitmap(self.awaitimage))
            val_change = 0
        elif text == 'rejected':
            self.res.SetBitmap(wx.Bitmap(self.crossimage))
            val_change = 2
        if self.dpt == "PE":
            self.PE_sign = val_change
        elif self.dpt == "QA":
            self.QA_sign = val_change
        elif self.dpt == "PD":
            self.PD_sign = val_change
        return

    def OnPopupItemSelected(self, event):
        item = self.popupmenu.FindItemById(event.GetId())
        text = item.GetText()

        if text == "view the PDF":
            try:
                pdf_name = wordy.extractedID[0] + '_' + wordy.extractedID[1] \
                    + '.pdf'
                (ignore, current_doc) = search_file(wordy.rootpath, pdf_name)
                if os.path.isfile(current_doc):
                    SP_Popen(current_doc, shell=True)
            except TypeError:
                print ("PDF is missing")
        elif text == "paste":
            if self.res == self.m_textCtrl_PEOI or self.res == self.m_textCtrl_ISS:
                if wordy.userlevel != 'master':
                    dlg = wx.MessageDialog(None, 'Insufficient permission to paste in to this field', '',
                                           wx.OK | wx.ICON_ERROR)
                    dlg.ShowModal()
                    dlg.Destroy()
                    return
            not_empty = wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_TEXT))
            if not_empty:
                text_data = wx.TextDataObject()
                if wx.TheClipboard.Open():
                    success = wx.TheClipboard.GetData(text_data)
                    wx.TheClipboard.Close()
                if success:
                    clipboardtext = safestr(text_data.GetText(), True)
                    self.res.SetLabel(clipboardtext)
        elif text == "today":
            if self.QA_sign == 1:
                if self.PD_sign == 1:
                    if self.PE_sign == 1:
                        self.m_textCtrl_dateRC.SetLabel(
                            safestr(datetime.datetime.today().strftime('%d-%m-%Y')))
        elif text == "me":
            config = configparser.ConfigParser()
            config.read('settings.ini')
            try:
                self.res.SetLabel(safestr(config['user_profile']['me']))
            except KeyError:
                pass
        elif text == "set me":
            set_me(self.panel)
        elif text == "Prod desc guess":
            self.go_guess(1)
        elif text == "Cust PN guess":
            self.go_guess(2)
        elif text == "Stages guess":
            self.go_guess(3)
        elif text == "Stage desc guess":
            self.go_guess(4)
        elif text == "Pek PN guess":
            self.go_guess(5)
        elif text =='guess all':
            self.m_textCtrl_cust_PN.SetLabel(wordy.headerstuff[2])
            self.m_textCtrl_pekPN.SetLabel(wordy.headerstuff[5])
            self.m_textCtrl_proc.SetLabel(wordy.headerstuff[4])
            self.m_textCtrl_prod_desc.SetLabel(wordy.headerstuff[1])
            self.m_textCtrl_stages.SetLabel(wordy.headerstuff[3])
        elif text == 'First Issue':
            self.m_textCtrl_DOC.SetLabel("First Issue")
            self.m_textCtrl_RFC.SetLabel("First Issue")
                    
    def go_guess(self, num):
        if wordy.headerstuff:
            if wordy.headerstuff[0] == wordy.extractedID[0]:    
                self.res.SetLabel(wordy.headerstuff[num])

    def on1Focus(self, event, res, pmu, dpt):
        self.res = res
        self.dpt = dpt
        pos = event.GetPosition()
        pos = self.panel.ScreenToClient(pos)
        self.panel.PopupMenu(pmu, pos)
        return

    def pages(self):
        if os.path.exists(wordy.file_out):
            with open(wordy.file_out, 'rb') as filehandle_input:

                # read content of the original file

                pdf = PdfFileReader(filehandle_input)
                x = 0
                NumPages = pdf.getNumPages()
                for i in range(0, NumPages):
                    PageObj = pdf.getPage(i)
                    Text = PageObj.extractText().split()
                    if len(Text)>1:
                        x+=1
            wordy.file_out = None
        else:
            pass
        self.m_textCtrl_pages.SetLabel(safestr(x))

    def load_data(self):
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            return

        skippy = False
        found_you = -1
        found2 = -1
        found_you = findme(self.roctop, wordy.extractedID[0])

        if wordy.do_save == 1:
            if wordy.file_out is not None:
                self.pages()
        self.PD_sign = 0
        self.PE_sign = 0
        self.QA_sign = 0
        if found_you >= 0:
            self.m_textCtrl_cust_PN.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                               ]['cstpn'], True))
            self.m_textCtrl_proc.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                            ]['stagedesc'], True))
            self.m_textCtrl_pekPN.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                             ]['pekpn'], True))
            self.m_textCtrl_prod_desc.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                                 ]['desc'], True))
            self.m_textCtrl_stages.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                              ]['stageno'], True))
            self.isobs.SetValue(self.roctop[found_you][0]['details'].get('obsolete', False))
            self.onChecked(None)
            if len(self.roctop[found_you]) > 0:
                findme2 = findingNemo(self.roctop[found_you],
                                      wordy.extractedID[1])
                found2 = findme2[0]
                skippy = findme2[1]
            if found2 > 0:
                if skippy is False:
                    self.m_textCtrl_DOC.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                        ]['doc'], True))
                    self.m_textCtrl_RFC.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                        ]['rfc'], True))
                    self.m_textCtrl_pages.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                          ]['pages'], True))
                    self.m_textCtrl_impl.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                         ]['impl'], True))
                    self.m_textCtrl_copies.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                           ]['copies'], True))
                    self.m_textCtrl_rcvd.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                         ]['rcvd'], True))
                    self.m_textCtrl_dateRC.SetLabel(safestr(self.roctop[found_you][found2]['revs'
                                                                                           ]['date'], True))

                    self.dotickboxes(found_you, found2)

                if found2 > 1:
                    pos = found2 - 5
                    acounter = found2 - 1
                    xr = 0
                    while acounter > found2 - 5:
                        if acounter > 0:
                            self.roc_prev.SetCellValue(xr, 0,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['rev'], True))
                            self.roc_prev.SetCellValue(xr, 1,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['doc'], True))
                            self.roc_prev.SetCellValue(xr, 2,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['rfc'], True))
                            self.roc_prev.SetCellValue(xr, 3,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['pages'], True))
                            self.roc_prev.SetCellValue(xr, 4,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['copies'], True))
                            self.roc_prev.SetCellValue(xr, 5,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['impl'], True))
                            self.roc_prev.SetCellValue(xr, 6,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['rcvd'], True))
                            self.roc_prev.SetCellValue(xr, 7,
                                                       safestr(self.roctop[found_you][acounter]['revs'
                                                                                                ]['date'], True))
                        xr += 1
                        acounter -= 1

        self.skippy = skippy
        self.found_you = found_you
        self.found2 = found2
        self.no_edits()

    def dotickboxes(self, found_you, found2):
        try:
            if self.roctop[found_you][found2]:
                pass
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)                      
            print("wtf is this error")
            print (self.roctop[found_you])
            return
        if self.roctop[found_you][found2]['revs']['PE_sign'
                                                  ] == 1:
            self.m_bitmap0.SetBitmap(wx.Bitmap(self.tickimage))
            self.PE_sign = 1
        elif self.roctop[found_you][found2]['revs'
                                            ]['PE_sign'] == 0:
            self.m_bitmap0.SetBitmap(wx.Bitmap(self.awaitimage))
            self.PE_sign = 0
        elif self.roctop[found_you][found2]['revs'
                                            ]['PE_sign'] == 2:
            self.m_bitmap0.SetBitmap(wx.Bitmap(self.crossimage))
            self.PE_sign = 2

        self.dosigscheck(
            'PE_sign_name', self.PE_sig_name, "PE: ", found_you, found2)
        self.dosigscheck(
            'QA_sign_name', self.QA_sig_name, "QA: ", found_you, found2)
        self.dosigscheck(
            'PD_sign_name', self.PD_sig_name, "PD: ", found_you, found2)

        if self.roctop[found_you][found2]['revs']['QA_sign'
                                                  ] == 1:
            self.m_bitmap1.SetBitmap(wx.Bitmap(self.tickimage))
            self.QA_sign = 1
        elif self.roctop[found_you][found2]['revs'
                                            ]['QA_sign'] == 0:
            self.m_bitmap1.SetBitmap(wx.Bitmap(self.awaitimage))
            self.QA_sign = 0
        elif self.roctop[found_you][found2]['revs'
                                            ]['QA_sign'] == 2:
            self.m_bitmap1.SetBitmap(wx.Bitmap(self.crossimage))
            self.QA_sign = 2

        if self.roctop[found_you][found2]['revs']['PD_sign'
                                                  ] == 1:
            self.m_bitmap2.SetBitmap(wx.Bitmap(self.tickimage))
            self.PD_sign = 1
        elif self.roctop[found_you][found2]['revs'
                                            ]['PD_sign'] == 0:
            self.m_bitmap2.SetBitmap(wx.Bitmap(self.awaitimage))
            self.PD_sign = 0
        elif self.roctop[found_you][found2]['revs'
                                            ]['PD_sign'] == 2:
            self.m_bitmap2.SetBitmap(wx.Bitmap(self.crossimage))
            self.PD_sign = 2

        # do the grey out or not check here
        self.m_textCtrl_rcvd.Disable()
        self.m_textCtrl_dateRC.Disable()
        if wordy.userlevel == "master":
            if self.QA_sign == 1:
                if self.PD_sign == 1:
                    if self.PE_sign == 1:
                        self.m_textCtrl_rcvd.Enable()
                        self.m_textCtrl_dateRC.Enable()
        # grey it out

    def reset_approvals(self, event):
        self.m_bitmap0.SetBitmap(wx.Bitmap(self.awaitimage))
        self.PE_sign = 0
        self.m_bitmap1.SetBitmap(wx.Bitmap(self.awaitimage))
        self.QA_sign = 0
        self.m_bitmap2.SetBitmap(wx.Bitmap(self.awaitimage))
        self.PD_sign = 0
        self.m_textCtrl_rcvd.Disable()
        self.m_textCtrl_dateRC.Disable()
        if wordy.userlevel == "master":
            if self.QA_sign == 1:
                if self.PD_sign == 1:
                    if self.PE_sign == 1:
                        self.m_textCtrl_rcvd.Enable()
                        self.m_textCtrl_dateRC.Enable()
        return

    def dosigscheck(self, x1, x2, x3, found_you, found2):
        signers = self.roctop[found_you][found2].get(
            'revs', {}).get(x1)
        if signers is not None:
            x2.SetLabel(x3 + signers)

    def submitter(self, event):
        skippy = False
        found_you = -1
        found2 = -1
        if not os.path.isfile(self.theDB):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return
        with open(self.theDB, 'rb') as f:
            roc2 = PKL_load(f)
        wordy.extractedID[1] = self.m_textCtrl_ISS.GetValue()
        wordy.extractedID[0] = self.m_textCtrl_PEOI.GetValue()
        found_you = findme(roc2, wordy.extractedID[0])
        if found_you >= 0:
            if len(roc2[found_you]) > 0:
                findme2 = findingNemo(roc2[found_you],
                                      wordy.extractedID[1])
                found2 = findme2[0]
                skippy = findme2[1]
        if self.skippy == skippy and self.found2 == found2 and self.found_you == found_you:
            pass
        else:
            if self.skippy != skippy:
                if self.skippy is True and skippy is False:
                    dlg = wx.MessageDialog(None,
                                           'Since change control window was opened the database has been updated and this revision level now exists. This window will close without any changes being saved', '', wx.OK | wx.ICON_ERROR)
                    dlg.ShowModal()
                    dlg.Destroy()
                    self.Close()
                    return
            if self.found2 != found2:
                pass
            if self.found_you != found_you:
                if self.found_you < 0 and found_you >= 0:
                    dlg = wx.MessageDialog(None,
                                           "since change control window was opened the database has been updated and 'details' data now exists. This window will close without any changes being saved", '', wx.OK | wx.ICON_ERROR)
                    dlg.ShowModal()
                    dlg.Destroy()
                    self.Close()
                    return
        self.roctop = roc2
        templist1 = [
            wordy.extractedID[0],
            safestr(self.m_textCtrl_cust_PN.GetValue()),
            safestr(self.m_textCtrl_pekPN.GetValue()),
            safestr(self.m_textCtrl_prod_desc.GetValue()),
            safestr(self.m_textCtrl_stages.GetValue()),
            safestr(self.m_textCtrl_proc.GetValue()),
            self.isobs.GetValue()
        ]
        templist2 = [
            wordy.extractedID[1],
            safestr(self.m_textCtrl_DOC.GetValue()),
            safestr(self.m_textCtrl_RFC.GetValue()),
            safestr(self.m_textCtrl_pages.GetValue()),
            safestr(self.m_textCtrl_copies.GetValue()),
            safestr(self.m_textCtrl_impl.GetValue()),
            safestr(self.m_textCtrl_rcvd.GetValue()),
            safestr(self.m_textCtrl_dateRC.GetValue()),
            self.PE_sign,
            self.QA_sign,
            self.PD_sign,
        ]

        # if self.found_you == -1:
        if found_you == -1:

            # new op.in and rev
            self.roctop.append(self.dict_dd(templist1, templist2))
            self.saverfunc()
            if wordy.do_save != 1:
                self.Close()
            else:
                wordy.do_save = 0
        # if self.found_you > 0 and self.skippy == True:
        if found_you > 0 and skippy is True:

            # op.in details updated, rev details all new

            self.roctop[found_you].append(self.addin_changes(templist2))
            self.roctop[found_you][0]['details'
                                      ].update(self.app_end_details(templist1))
            self.saverfunc()
            if wordy.do_save != 1:
                self.Close()
            else:
                wordy.do_save = 0
        # if self.found_you > 0 and self.skippy == False:
        if found_you > 0 and skippy is False:

            # op.in and rev details both updated

            self.roctop[found_you][found2]['revs'
                                           ].update(self.app_end_changes(templist2))
            self.roctop[found_you][0]['details'
                                      ].update(self.app_end_details(templist1))
            self.saverfunc()
            if wordy.do_save != 1:
                self.Close()
            else:
                wordy.do_save = 0
            if wordy.userlevel == 'master':
                self.Close()
                title = 'Record of Change to Operator Instructions'
                frame = OtherFrame(title=title)
                wordy.frame_number = 1

            return
        if wordy.userlevel == 'master':
            self.Close()
            title = 'Record of Change to Operator Instructions'
            frame = OtherFrame(title=title)
            wordy.frame_number = 1
        return

    def saverfunc(self):
        if not os.path.isfile(self.theDB):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return

        with open(self.theDB, 'wb') as f:
            PKL_dump(self.roctop, f, indent = 2)
        print(datetime.datetime.now())
        if wordy.do_save != 1:
            dlg = wx.MessageDialog(None,
                                   'Record of Change Database updated',
                                   '', wx.OK)
            dlg.ShowModal()
            dlg.Destroy()
        return

    def dict_dd(self, detaildata, changes):
        detaildict = {
            'opno': detaildata[0],
            'cstpn': detaildata[1],
            'pekpn': detaildata[2],
            'desc': detaildata[3],
            'stageno': detaildata[4],
            'stagedesc': detaildata[5],
            'obsolete': detaildata[6],
        }
        useradded2 = {'details': detaildict}
        useradded = self.addin_changes(changes)
        info = [useradded2, useradded]
        return info

    def addin_changes(self, changes):
        try:
            new_extra = wordy.extractedID[2]
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            new_extra = None

        date_entered = time()

        changesdict = {
            'rev': changes[0],
            'doc': changes[1],
            'rfc': changes[2],
            'pages': changes[3],
            'copies': changes[4],
            'impl': changes[5],
            'rcvd': changes[6],
            'date': changes[7],
            'PE_sign': changes[8],
            'QA_sign': changes[9],
            'PD_sign': changes[10],
            'rev_iss_date': new_extra,
            'QA_sign_name': None,
            'PE_sign_name': None,
            'PD_sign_name': None,
            'date_entered': date_entered,
        }
        useradded = {'revs': changesdict}
        return useradded

    def app_end_changes(self, changes):
        changesdict = {
            'rev': changes[0],
            'doc': changes[1],
            'rfc': changes[2],
            'pages': changes[3],
            'copies': changes[4],
            'impl': changes[5],
            'rcvd': changes[6],
            'date': changes[7],
            'PE_sign': changes[8],
            'QA_sign': changes[9],
            'PD_sign': changes[10],
        }
        return changesdict

    def app_end_details(self, detaildata):
        detaildict = {
            'opno': detaildata[0],
            'cstpn': detaildata[1],
            'pekpn': detaildata[2],
            'desc': detaildata[3],
            'stageno': detaildata[4],
            'stagedesc': detaildata[5],
            'obsolete': detaildata[6],
        }
        return detaildict

    def __del__(self):
        wordy.frame_number = 0


class wordy(wx.Frame):

    def __init__(self, parent, title):
        super(wordy, self).__init__(
            parent,
            id=wx.ID_ANY,
            title=title,
            pos=wx.DefaultPosition,
            size=wx.Size(530, 335),
            style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU |
            wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN | wx.TAB_TRAVERSAL,
        )
        try:
            if wordy.suppress == 1:
                pass
        except AttributeError:
            wordy.suppress = 0

        if wordy.suppress == 0:
            bitmap = wx.Bitmap(SPLASHER.getBitmap())
            splash = wx.adv.SplashScreen(
                bitmap, 
                wx.adv.SPLASH_CENTER_ON_SCREEN|wx.adv.SPLASH_TIMEOUT, 
                3000, self)
        
            splash.Show()      
            self.SetIcon(OICC_LOGO.GetIcon())
            config = configparser.ConfigParser()
            try:
                config.read('settings.ini')
                paff = config['rootpath']['path']
                self.rootpath = paff + '\\'
                wordy.userlevel = config['user_profile']['department']
                os.system('mode con: cols=%s lines=%s'
                          % (config['consolesize']['cols'],
                             config['consolesize']['lines']))
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                self.rootpath = r'\\NT4\Client_Files\Public\PEOI' + '\\'
                wordy.userlevel = 'none'
                os.system('mode con: cols=15 lines=1')
                pass
            try:
                if len(config['user_profile']['me']) < 1:
                    set_me(None)
            except KeyError:
                set_me(None)
            self.SetFont(wx.Font(
                9,
                wx.MODERN,
                wx.NORMAL,
                wx.NORMAL,
                False,
                u'Consolas',
            ))
            wordy.rootpath = self.rootpath
            path4approval = r'for_approval' + '\\'
            pathApproved = r'approved' + '\\'
            pathArchive = r'archive' + '\\'
            pathRejected = r'rejected' + '\\'
            pathROC = r'ROC' + '\\'
            pathDB = r'DB' + '\\'
            pathDBarch = r'arch' + '\\'
            DBfn = r'ROC_db.json'
            wordy.temp_directory = os.path.join(gettempdir(), '.{}'.format(hash(os.times())))
            try:
                os.makedirs(wordy.temp_directory)
            except WindowsError:
                pass
            temper = safestr(map(ord, os.urandom(1))[0]) + r"temp.docx"
            temper2 = safestr(map(ord, os.urandom(1))[0]) + r"temp.doc"
            temPDFer = safestr(map(ord, os.urandom(1))[0]) + r"temp_ony.pdf"
            self.ROCpath = os.path.abspath(self.rootpath
                                           + pathROC)
            self.forapprovalpath = os.path.abspath(self.rootpath
                                                   + path4approval)
            wordy.forappralpath = self.forapprovalpath
            self.approvedpath = os.path.abspath(self.rootpath
                                                + pathApproved)
            wordy.approvedpath = self.approvedpath
            self.transition_file = os.path.abspath(self.rootpath
                                                   + temper)
            self.transition_file2 = os.path.abspath(self.rootpath
                                                    + temper2)
            self.output_file = os.path.abspath(self.rootpath + temPDFer)
            self.archive_path = os.path.abspath(self.rootpath
                                                + pathArchive)
            wordy.archive_path = self.archive_path
            self.reject_path = os.path.abspath(self.rootpath
                                               + pathRejected)
            wordy.DBpath = os.path.abspath(self.rootpath + pathDB
                                           + DBfn)
            self.DBpathARCH = os.path.abspath(self.rootpath + pathDBarch)
            self.databasecsv = r"\\Nt4\Client_Files\Public\Personnel\TRAINING_RECORDS\opcards\__SUMMARY.txt"
            wordy.loggingDB = os.path.abspath(
                wordy.rootpath + '\\' + 'finalised' + '\\' + 'logging.db')
            wordy.loggingrecent = os.path.abspath(
                wordy.rootpath + '\\' + 'finalised' + '\\' + 'updated.txt')
            wordy.email = os.path.abspath(
                wordy.rootpath + '\\' + 'PEOIs_released.msg')
            wordy.file_out = None
            wordy.reset_after_reimport = 0
            bSizer1 = wx.BoxSizer(wx.VERTICAL)
            bSizer1.SetMinSize(wx.Size(-1, 55))
            bSizer2 = wx.BoxSizer(wx.HORIZONTAL)
            wordy.output_file = self.output_file
            self.m_button4 = wx.Button(
                self,
                wx.ID_ANY,
                u'import original',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer2.Add(self.m_button4, 1, wx.ALL, 5)

            self.m_button5 = wx.Button(
                self,
                wx.ID_ANY,
                u'PE approve',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer2.Add(self.m_button5, 1, wx.ALL, 5)

            self.m_button7 = wx.Button(
                self,
                wx.ID_ANY,
                u'QA review',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer2.Add(self.m_button7, 1, wx.ALL, 5)

            self.m_button6 = wx.Button(
                self,
                wx.ID_ANY,
                u'Prod review',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer2.Add(self.m_button6, 1, wx.ALL, 5)

            bSizer1.Add(bSizer2, 1, wx.EXPAND, 5)

            bSizer23 = wx.BoxSizer(wx.HORIZONTAL)
            bSizerX23X = wx.BoxSizer(wx.VERTICAL)

            bSizerX23X2 = wx.StaticBoxSizer(wx.StaticBox(self,
                                                         wx.ID_ANY, wx.EmptyString), wx.HORIZONTAL)
            bSizerX23X.SetMinSize(wx.Size(235, 40))
            bSizerX23X2.SetMinSize(wx.Size(235, 40))

            self.m_button41 = wx.Button(
                self,
                wx.ID_ANY,
                u'submit imported \n to database',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer23.Add(self.m_button41, 0, wx.ALL | wx.EXPAND
                         | wx.CENTER, 5)

            self.m_staticText242 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'Save File Name:',
                wx.DefaultPosition,
                wx.Size(105, -1),
                style=wx.ALIGN_RIGHT | wx.TE_MULTILINE,
            )
            self.m_staticText242.SetMinSize(wx.Size(105, -1))
            self.m_staticText242.Wrap(-1)
            bSizer23.Add(self.m_staticText242, 1, wx.TOP | wx.EXPAND,
                         35)
            self.m_staticlinexx1 = wx.StaticLine(self, wx.ID_ANY,
                                                 wx.DefaultPosition, wx.DefaultSize,
                                                 wx.LI_HORIZONTAL)

            bSizerX23X.Add(self.m_staticlinexx1, 0, wx.ALL, 5)
            self.m_staticText241 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'-no file-',
                wx.DefaultPosition,
                wx.Size(195, -1),
                style=wx.ALIGN_CENTER | wx.TE_MULTILINE,
            )
            self.m_staticText241.SetMinSize(wx.Size(195, -1))
            self.m_staticText241.Wrap(-1)
            bSizerX23X.Add(self.m_staticText241, 1, wx.EXPAND
                           | wx.CENTER | wx.FIXED_MINSIZE
                           | wx.RESERVE_SPACE_EVEN_IF_HIDDEN, 0)
            bSizerX23X2.Add(bSizerX23X, 1, wx.EXPAND | wx.CENTER
                            | wx.FIXED_MINSIZE
                            | wx.RESERVE_SPACE_EVEN_IF_HIDDEN, 0)
            bSizer23.Add(bSizerX23X2, 1, wx.ALL | wx.EXPAND
                         | wx.RESERVE_SPACE_EVEN_IF_HIDDEN, 10)

            bSizer1.Add(bSizer23, 1, wx.EXPAND, 5)

            bSizer21 = wx.BoxSizer(wx.HORIZONTAL)

            bSizer3 = wx.BoxSizer(wx.VERTICAL)

            bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

            self.m_textCtrl4 = wx.TextCtrl(
                self,
                wx.ID_ANY,
                u'blank',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )

            self.m_staticText2 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'PEOI ',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.m_staticText2.Wrap(-1)
            bSizer10.Add(self.m_staticText2, 0, wx.ALL, 5)
            bSizer10.Add(self.m_textCtrl4, 1, wx.ALL, 5)

            bSizer3.Add(bSizer10, 1, wx.EXPAND, 5)

            bSizer101 = wx.BoxSizer(wx.HORIZONTAL)

            self.m_textCtrl42 = wx.TextCtrl(
                self,
                wx.ID_ANY,
                u'blank',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.m_staticText21 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'ISSUE',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.m_staticText21.Wrap(-1)
            bSizer101.Add(self.m_staticText21, 0, wx.ALL, 5)
            bSizer101.Add(self.m_textCtrl42, 1, wx.ALL, 5)

            bSizer3.Add(bSizer101, 1, wx.EXPAND, 5)

            bSizer102 = wx.BoxSizer(wx.HORIZONTAL)

            self.m_textCtrl43 = wx.TextCtrl(
                self,
                wx.ID_ANY,
                u'blank',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.m_staticText22 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'DATE ',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.m_staticText22.Wrap(-1)
            bSizer102.Add(self.m_staticText22, 0, wx.ALL, 5)
            bSizer102.Add(self.m_textCtrl43, 1, wx.ALL, 5)

            bSizer3.Add(bSizer102, 1, wx.EXPAND, 5)

            bSizer21.Add(bSizer3, 1, wx.EXPAND, 5)

            bSizer31 = wx.BoxSizer(wx.VERTICAL)

            bSizer103 = wx.BoxSizer(wx.HORIZONTAL)

            bSizer31.Add(bSizer103, 1, wx.EXPAND, 5)

            bSizer1041 = wx.BoxSizer(wx.VERTICAL)

            self.backUPtheDB()

            self.m_button712 = wx.Button(
                self,
                wx.ID_ANY,
                u'Open Record of change doc for:',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizer1041.Add(self.m_button712, 0, wx.ALL
                           | wx.ALIGN_CENTER_VERTICAL, 5)

            self.m_comboBox2ChoicesFULL = self.poplist()

            self.m_comboBox2 = wx.ComboBox(
                self,
                wx.ID_ANY,
                wx.EmptyString,
                wx.DefaultPosition,
                wx.DefaultSize,
                self.m_comboBox2ChoicesFULL,
                0,
            )
            bSizer1041.Add(self.m_comboBox2, 0, wx.ALL
                           | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

            bSizer31.Add(bSizer1041, 1, wx.EXPAND, 5)

            bSizer21.Add(bSizer31, 1, wx.EXPAND, 5)

            bSizer1.Add(bSizer21, 1, wx.EXPAND, 5)

            bSizer22 = wx.BoxSizer(wx.HORIZONTAL)

            self.m_button42 = wx.Button(
                self,
                wx.ID_ANY,
                u'FINALISE',
                wx.DefaultPosition,
                wx.Size(100, -1),
                0,
            )
            bSizer22.Add(self.m_button42, 0, wx.ALL, 5)

            self.m_button51 = wx.Button(
                self,
                wx.ID_ANY,
                u"Create 'Recently Updated' Report",
                wx.DefaultPosition,
                wx.Size(230, -1),
                0,
            )
            bSizer22.Add(self.m_button51, 0, wx.ALL, 5)

            self.m_staticText2x2 = wx.StaticText(
                self,
                wx.ID_ANY,
                u'Select a file to Archive: ',
                wx.DefaultPosition,
                wx.Size(100, 30),
                0,
            )
            self.m_staticText2x2.Wrap(-1)
            bSizer22.Add(self.m_staticText2x2, 0, wx.ALL
                         | wx.ALIGN_RIGHT, 5)

            self.m_filePicker1 = wx.FilePickerCtrl(
                self,
                wx.ID_ANY,
                wx.EmptyString,
                u'Archive a file',
                u'*.pdf',
                wx.DefaultPosition,
                wx.DefaultSize,
                wx.FLP_SMALL,
            )
            self.m_filePicker1.SetInitialDirectory(self.approvedpath)
            bSizer22.Add(self.m_filePicker1, 0, wx.ALL, 5)

            self.m_filePicker1.Bind(wx.EVT_FILEPICKER_CHANGED,
                                    self.m_filePicker1OnFileChanged)

            bSizer1.Add(bSizer22, 1, wx.EXPAND, 5)

            bSizerQQ = wx.BoxSizer(wx.HORIZONTAL)
            self.m_buttonq42 = wx.Button(
                self,
                wx.ID_ANY,
                u'Read-Only R.O.C open',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            self.sortorder = wx.CheckBox(
                self,
                wx.ID_ANY,
                u"",
                wx.DefaultPosition,
                wx.DefaultSize,
                0,                       
            )
            (self.m_comboBoxDB, self.m_comboBoxDBTL) = self.DBpoplist()

            bSizerQQ.Add(self.sortorder, 0, wx.ALL
                         | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

            self.m_comboBox2TL = wx.ComboBox(
                self,
                wx.ID_ANY,
                wx.EmptyString,
                wx.DefaultPosition,
                wx.DefaultSize,
                self.m_comboBoxDBTL,
                0,
            )
            bSizerQQ.Add(self.m_comboBox2TL, 0, wx.ALL
                         | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)
            self.m_comboBox2qq = wx.ComboBox(
                self,
                wx.ID_ANY,
                wx.EmptyString,
                wx.DefaultPosition,
                wx.DefaultSize,
                self.m_comboBoxDB,
                0,
            )
            bSizerQQ.Add(self.m_comboBox2qq, 0, wx.ALL
                         | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

            bSizerQQ.Add(self.m_buttonq42, 0, wx.ALL, 5)
            if wordy.userlevel == "master":

                godframe = GodMode_Frame("Warning, God Mode Enabled")

            self.m_buttonq42.SetLabel('R.O. open PEOI')
            self.m_buttonMASTER = wx.Button(
                self,
                wx.ID_ANY,
                u'Open ISS',
                wx.DefaultPosition,
                wx.DefaultSize,
                0,
            )
            bSizerQQ.Add(self.m_buttonMASTER, 0, wx.ALL, 5)
            self.tlist3 = ['']
            self.m_rev_master = wx.ComboBox(
                self,
                wx.ID_ANY,
                wx.EmptyString,
                wx.DefaultPosition,
                wx.Size(80, -1),
                self.tlist3,
                0,
            )
            bSizerQQ.Add(self.m_rev_master, 0, wx.ALL, 5)

            bSizer1.Add(bSizerQQ, 1, wx.EXPAND, 5)
            self.Layout()
            menubar = wx.MenuBar()
            fileMenu = wx.Menu()
            file1Menu1 = wx.Menu()
            file2Menu2 = wx.Menu()
            file3Menu3 = wx.Menu()
            file4Menu4 = wx.Menu()
            fileMenu.Append(1, 'Prod Eng user manual', 'Help Files')
            file1Menu1.Append(2, 'QA user manual', 'Help Files')
            file3Menu3.Append(
                3, 'Production Managers user manual', 'Help Files')
            fileMenu.Append(4, 'Prod Eng: rectifying rejection', 'Help Files')
            file2Menu2.Append(
                5, 'All: sign-off setup in Cute PDF', 'Help Files')
            file2Menu2.Append(6, 'PE, PD: finalisation', 'Help Files')
            file2Menu2.Append(7, 'other stuff', 'Help Files')
            file4Menu4.Append(8, 'PEOIs Outstanding', 'Outstanding')
            file4Menu4.Append(
                9, 'Update Printable Records of Change', 'Print ROCS up')
            file4Menu4.Append(
                10, 'update technical after finalisation', 'update tech')
            file4Menu4.Append(
                11, 'View Printable Records of Change', 'ROCS view')
            file4Menu4.Append(
                12, 'Create Operator Index', 'OPINDEX')
            file4Menu4.Append(
                13, 'Create Detailed Outstanding List', 'detailed outstanding list')
            file4Menu4.Append(
                14, 'Generate Recently Updated Report', 'Recently Updated')
            file4Menu4.Append(
                15, 'View state of in progress files', 'In Progress Report')
            menubar.Insert(0, file2Menu2, '& General Help')
            menubar.Insert(1, fileMenu, '& Prod Eng Help')
            menubar.Insert(2, file1Menu1, '& QA Help')
            menubar.Insert(3, file3Menu3, '& Prod Dpt Help')
            menubar.Insert(4, file4Menu4, '& Reporting Tools')

            self.SetMenuBar(menubar)

            self.Bind(wx.EVT_MENU, self.helpfiles)
            self.SetSizer(bSizer1)

            self.Centre(wx.BOTH)
            self.Show()

            self.PEOI = ''
            self.ISSUE = ''
            self.DATE = ''
            self.extracted = [self.PEOI, self.ISSUE, self.DATE]
            wordy.extractedID = self.extracted
            wordy.reportfunc = 0
            self.sortorder.Bind(wx.EVT_CHECKBOX, self.update_the_ROC_list)
            self.m_button4.Bind(wx.EVT_BUTTON, self.import_original)
            self.m_button41.Bind(wx.EVT_BUTTON, self.in2PDF)
            self.m_textCtrl4.Bind(wx.EVT_TEXT, self.textupdated)
            self.m_textCtrl42.Bind(wx.EVT_TEXT, self.textupdated)
            self.m_textCtrl43.Bind(wx.EVT_TEXT, self.textupdated)
            self.m_button712.Bind(wx.EVT_BUTTON, self.openROC)
            self.m_button51.Bind(wx.EVT_BUTTON, self.loggindb)
            self.m_button7.Bind(wx.EVT_BUTTON, lambda event:
                                self.approve_frame(event, 'QA'))
            self.m_button5.Bind(wx.EVT_BUTTON, lambda event:
                                self.approve_frame(event, 'PE'))
            self.m_button6.Bind(wx.EVT_BUTTON, lambda event:
                                self.approve_frame(event, 'PD'))
            self.m_button42.Bind(wx.EVT_BUTTON, self.fin)
            self.m_buttonq42.Bind(wx.EVT_BUTTON, self.RO_ROC)
            self.m_buttonMASTER.Bind(wx.EVT_BUTTON, self.God_Mode)
            self.m_comboBox2TL.Bind(wx.EVT_COMBOBOX, self.filter_projs)
            self.m_comboBox2qq.Bind(wx.EVT_COMBOBOX, self.filter_revs)
           
            if wordy.userlevel == 'master':
                pass
            elif wordy.userlevel == 'none':

                # read only, view reports

                self.m_button7.Disable()
                self.m_button5.Disable()
                self.m_button6.Disable()
                self.m_button712.Disable()
                self.m_button42.Disable()
                self.m_button4.Disable()
                self.m_button41.Disable()
                self.m_filePicker1.Disable()
            elif wordy.userlevel == 'QA':

                # qa restrictions

                self.m_button5.Disable()
                self.m_button6.Disable()
                self.m_button42.Disable()
                self.m_button4.Disable()
                self.m_button41.Disable()
                self.m_filePicker1.Disable()
            elif wordy.userlevel == 'PE':

                # prod eng restrictions

                self.m_button7.Disable()
                self.m_button6.Disable()
            elif wordy.userlevel == 'PD':

                # prod dept restrictions

                self.m_button7.Disable()
                self.m_button5.Disable()
                self.m_button4.Disable()
                self.m_button41.Disable()
            else:

                # read only, view reports

                self.m_button7.Disable()
                self.m_button5.Disable()
                self.m_button6.Disable()
                self.m_button712.Disable()
                self.m_button42.Disable()
                self.m_button4.Disable()
                self.m_button41.Disable()
                self.m_filePicker1.Disable()

            wordy.frame_number = 0
            wordy.app_frame_number = 0
            wordy.final_frame_number = 0
            wordy.reportfunc = 2
            wordy.ROROC_frame_number = 0
            wordy.headerstuff = False
            FinaliseFrame('nothing')
            self.update_the_ROC_list()

        else:
            self.Close()

    def loggindb(self, event=None):
        if 1 == 1:
            # ask if they are opening to email report, or just to look

            del_or_no = wx.MessageDialog(None,
                                         ('Do you just wish to view the list of recently updated op-ins or distribute the list?') + '\n' +
                                         ('If you click YES, you must copy and paste the text file in to an email and distribute it.') + '\n' +
                                         ('CLick YES if you are going to do this or NO if you just want to view the report.'),
                                         ('Just view report = NO, Distribute via email = YES'),
                                         wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION).ShowModal()
            if del_or_no != wx.ID_YES:
                do_delete = False
            elif del_or_no == wx.ID_YES:
                do_delete = True
            connection_obj = sqlite3.connect(wordy.loggingDB)
            # wordy.loggingrecent
            # cursor object
            # get everything from DB and write it to text file
            outlist = []
            # this gets the DB stuff
            cursor_obj = connection_obj.cursor()
            # connection object - updated_instructions is the 'table'
            cursor_obj.execute("SELECT * FROM updated_instructions")
            # get everything
            x = (cursor_obj.fetchall())
            # get it all in one simple list
            if len(x) > 0:
                for p in x:
                    for j in p:
                        outlist.append(j)
                    outlist.append("")
                    outlist.append("---")
                    outlist.append("")
                # write to text file
                with open(wordy.loggingrecent, 'w') as f:
                    for item in outlist:
                        f.write("%s\n" % item)

                if do_delete:
                    # delete data
                    cursor_obj.execute("DELETE FROM updated_instructions")
                    connection_obj.commit()
                    # Close the connection
                    connection_obj.close()

                    wait = wx.BusyInfo(
                        "Copy & paste text file contents in to an email to report recently updated instructions")
                    sleep(1.3)

                    if os.path.isfile(wordy.email):
                        SP_Popen([wordy.email], shell=True)
                    else:
                        print(wordy.email)
                else:
                    wait = wx.BusyInfo(
                        "These are all instructions updated since the last report was distributed" + '\n' +
                        "You are just viewing this file and have not altered the contents of the databse.")
                    sleep(1.3)
            else:
                wait = wx.BusyInfo(
                    "No New updates were found since this was last run. Now opening the previous report" + '\n' +
                    "... the database indicates these have already been distributed in an email")
                sleep(1.3)
                if do_delete:
                    if os.path.isfile(wordy.email):
                        SP_Popen([wordy.email], shell=True)
                    else:
                        print(wordy.email)

            if os.path.isfile(wordy.loggingrecent):
                SP_Popen([wordy.loggingrecent], shell=True)
            else:
                pass
            del wait
        else:
            pass

    def helpfiles(self, event):
        evt_id = event.GetId()
        options = ['nothing', r"C:\Program Files (x86)\OICC\help\OICC_UM_PE.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_QA.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_PD.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_rectify_rejection.pdf",
                   r"C:\Program Files (x86)\OICC\help\OICC_UM_sign-off-setup.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_finalisation.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_other_stuff.pdf"]
        options2 = ['nothing', os.path.abspath(self.rootpath + r"\OICC_UM_PE.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_PE.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_PD.pdf"), os.path.abspath(
            self.rootpath + r"\OICC_UM_rectify_rejection.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_sign-off-setup.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_finalisation.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_other_stuff.pdf")]
        if evt_id < 8:
            if os.path.isfile(options[evt_id]):
                SP_Popen([options[evt_id]], shell=True)
            elif os.path.isfile(options[evt_id]):
                SP_Popen([options2[evt_id]], shell=True)
        else:
            if evt_id == 8:
                self.poplist()
                outs = os.path.abspath(self.rootpath + r'outstanding.html')
                if os.path.isfile(outs):
                    SP_Popen([outs], shell=True)
            elif evt_id == 9:
                rox = os.path.abspath(self.rootpath + r'\ROC')
                if os.path.exists(rox):
                    path = os.path.realpath(rox)
                    msg = wx.BusyInfo(
                        'Please wait, this will take some time to generate the files')
                    self.makeXML()
                    outs = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                pass
            elif evt_id == 10:
                if wordy.approvedpath == r'\\NT4\Client_Files\Public\PEOI\approved':
                    msg = wx.BusyInfo('Please wait, this will take some time')
                    self.xferttech(
                        r'\\AD2\Client_Files\Technical\PEOI\approved', wordy.approvedpath)
                else:
                    pass

            elif evt_id == 11:
                self.ROCxTOCxHTML()
                outs = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
                if os.path.isfile(outs):
                    SP_Popen([outs], shell=True)
                else:
                    pass

            elif evt_id == 12:
                outs = os.path.abspath(self.ROCpath + '\\' + "OPINDEX.csv")
                try:
                    with open(outs, 'wb') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                        wr.writerows(self.opindexgen())
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                    else:
                        pass
                except IOError:
                    print("cannot create opindex, someone has it open")
                    return

            elif evt_id == 13:
                outs = os.path.abspath(
                    self.ROCpath + '\\' + "detailed_outstanding.csv")
                try:
                    with open(outs, 'wb') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                        try:
                            x = self.det_out()
                            for y in x:
                                wr.writerows(y)
                        except Exception as e:
                            print('oops ' + str((inspect.stack()[0][2])))
                            print (e.message, e.args)
                            pass
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                    else:
                        pass
                except IOError:
                    print(
                        "cannot create detailed outstanding list, someone has it open")
                    return
            elif evt_id == 14:
                self.loggindb()
            elif evt_id == 15:
                self.reporter(event)
        return

    def XtracTor(self, target, pos):
        # it's the Vlookup again, bit fixed
        namez2 = []
        for nam3s in self.dbdata:
            temp = nam3s[0] + nam3s[1]
            if temp == target:
                temp2 = nam3s[pos]
                namez2.append(temp2)
        return namez2

    def RO_ROC(self, event):
        x = self.m_comboBox2qq.GetValue()
        if x != '':
            if x in self.m_comboBoxDB:
                if wordy.ROROC_frame_number == 0:
                    title = safestr(x)
                    frame = ROC_frame(title=title)
                    wordy.ROROC_frame_number = 1

    def God_Mode(self, event):
        x = self.m_comboBox2qq.GetValue()
        x2 = self.m_rev_master.GetValue()
        x3 = self.m_comboBox2TL.GetValue()
        if x != '' and x2 != '':
            if x in self.m_comboBoxDB:
                if x2 in self.tlist3:
                    if wordy.frame_number == 0:
                        wordy.extractedID = [safestr(x), safestr(x2),
                                             None]
                        wordy.do_save = 0
                        title = 'Record of Change God Mode'
                        frame = OtherFrame(title=title)
                        wordy.frame_number = 1
                        return
        elif x == '' and x2 == '' and x3 == '':
            if wordy.frame_number == 0:
                wordy.extractedID = ['', '',
                                     None]
                wordy.do_save = 0
                title = 'Record of Change God Mode'
                frame = OtherFrame(title=title)
                wordy.frame_number = 1

    def DBOK(self):
        if not os.path.isfile(wordy.DBpath):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return False
        else:
            return True

    def getnumfils(self):
        return len([name for name in os.listdir(self.DBpathARCH) if os.path.isfile(os.path.join(self.DBpathARCH, name))])

    def delOldest(self):
        oldest = min(os.listdir(self.DBpathARCH), key=lambda f: os.path.getctime(
            "{}/{}".format(self.DBpathARCH, f)))
        if fileaccesscheck(os.path.abspath(self.DBpathARCH + '\\' + oldest)):
            os.remove(os.path.abspath(self.DBpathARCH + '\\' + oldest))
        return

    def backUPtheDB(self):
        newest = max(os.listdir(self.DBpathARCH), key=lambda f: os.path.getctime(
            "{}/{}".format(self.DBpathARCH, f)))
        if (filecmp.cmp(wordy.DBpath, self.DBpathARCH + '\\' + newest, shallow=1)):
            return
        dumblist = [n for n in range(1, 99)]
        if self.getnumfils() > 0:
            for (dirpath, dirnames, fi) in os.walk(self.DBpathARCH):
                archfiles = []
                for fp in fi:
                    archfiles.append(
                        int(fp[-(len(fp) - (fp.rfind('.'))) + 1:]))
            newfile = 'ROC_db.pkl.' + \
                safestr(returnNotMatches(dumblist, archfiles)[0])
        else:
            newfile = 'ROC_db.pkl.1'
        try:
            shutil.copy(wordy.DBpath, os.path.abspath(
                self.DBpathARCH + '\\' + newfile))
            if self.getnumfils() > 10:
                while self.getnumfils() > 25:
                    if self.getnumfils() > 10:
                        self.delOldest()
                    else:
                        break
                rd = Path(self.DBpathARCH)
                while sum(f.stat().st_size for f in rd.glob('**/*') if f.is_file()) > 15000000:
                    if self.getnumfils() > 10:
                        self.delOldest()
                    else:
                        break
            else:
                return
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        return

    def det_out(self):
        valz = self.poplist(True)
        awQA = []
        awPE = []
        awPD = []
        rejd = []
        approved = []
        for x in valz:
            if isinstance(x, dict):
                PEstr = x.get('PE')
                PDstr = x.get('PD')
                QAstr = x.get('QA')
                RRstr = x.get('reject_reason')
                revstr = x.get('rev')
                PEOIstr = x.get('PEOI')
                DoCstr = x.get('doc')
                RfCstr = x.get('rfc')
                if isinstance(RRstr, list):
                    tt = safestr(''.join(RRstr))
                    RRstr = tt
                l1 = [PEstr, PDstr, QAstr]
                l2 = [awPE, awPD, awQA]
                for x, y in zip(l1, l2):
                    if x == 'rejected':
                        rejd.append(
                            [PEOIstr, revstr, DoCstr, RfCstr, RRstr])
                    elif x == 'awaiting':
                        if PEstr != 'rejected' and PDstr != 'rejected' and QAstr != 'rejected':
                            y.append(
                                [PEOIstr, revstr, DoCstr, RfCstr, RRstr])
                if PEstr == 'approved' and PDstr == 'approved' and QAstr == 'approved':
                    approved.append([PEOIstr, revstr, DoCstr, RfCstr, RRstr])
        inject = ["PEOI", "REV", "details of change",
                  "reason for change", "Comment"]
        inject2 = ["*****************", "*****",
                   "*****************", "*****************", "*****************"]

        valz4 = [
            [["No. in system: ", len(valz)]],
            [["Outstanding PE: ", len(awPE)]],
            [["Outstanding QA: ", len(awQA)]],
            [["Outstanding PD: ", len(awPD)]],
            [["Rejected: ", len(rejd)]],
            [["All Approved: ", len(approved)]],
            [inject2, ["Instructions Rejected"], [""], inject],
            rejd,
            [inject2, ["Instructions awaiting Prod Eng approval"], [""],  inject],
            awPE,
            [inject2, ["Instructions awaiting Production Manager approval"], [""],  inject],
            awPD,
            [inject2, ["Instructions awaiting Quality approval"], [""],  inject],
            awQA,
            [inject2, ["Instructions require issuing / finalising"], [""],  inject],
            approved
        ]

        return (valz4)

    def opindexgen(self):
        if self.DBOK():
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
        else:
            return
        PEOI = []
        for x in self.roctop:
            if isinstance(x, list):
                for y in x:
                    if isinstance(y, dict):
                        t = len(x)
                        if t > 1:
                            if isinstance(x[0], dict) and isinstance(x[-1], dict):
                                p1 = (x[0].get('details', {}).get('opno'))
                                p3 = (x[0].get('details', {}).get('obsolete', False))
                                p2 = (x[-1].get('revs', {}).get('rev'))
                                if p1 is not None and p2 is not None:
                                    PEOI.append(p1)
                                break
        PEOIx = sorted(PEOI)
        biglist = []
        biglist.append(["Project", "PEOI #", "Revision", "Description", "Implemented by",
                        "Date of revision*", "QA signed*", "Production signed*", "Date Received", "Details of Change", "Reason for Change", "Received by", "pages", "copies", "stage #s", "Pektron part numbers", "customer info", "link to file"])
        biglist.append(["*may not be availble (new feature May 2021)*"])
        for x in PEOIx:
            actpos = PEOI.index(x)
            opno = self.roctop[actpos][0].get('details', {}).get('opno')
            if x == opno:
                posy = len(self.roctop[actpos])
                if posy >= 1:
                    bl = self.verify_complete(actpos, posy - 1)
                    if bl:
                        biglist.append(bl)
        return(biglist)

    def verify_complete(self, actpos, cnt):
        opno = self.roctop[actpos][0].get('details', {}).get('opno')
        rev = self.roctop[actpos][cnt].get('revs', {}).get('rev')
        obstatus = self.roctop[actpos][0].get('details', {}).get('obsolete', False)
        if obstatus:
            obsval=" (PEOI marked obsolete)"
        else:
            obsval=""
        rev=rev+obsval
        opno=opno+obsval
        if int(self.roctop[actpos][cnt].get('revs', {}).get('PE_sign')) != 1:
            sig1 = 0
        else:
            sig1 = 1
        if int(self.roctop[actpos][cnt].get('revs', {}).get('PD_sign')) != 1:
            sig2 = 0
        else:
            sig2 = 1
        if int(self.roctop[actpos][cnt].get('revs', {}).get('QA_sign')) != 1:
            sig3 = 0
        else:
            sig3 = 1
        if sig1 + sig2 + sig3 == 3:
            impl = self.roctop[actpos][cnt].get('revs', {}).get('impl')
            desc = self.roctop[actpos][0].get('details', {}).get('stagedesc')
            OPINDEX_ISS = self.roctop[actpos][cnt].get(
                'revs', {}).get('rev_iss_date')
            OPINDEX_IMPL2 = self.roctop[actpos][cnt].get(
                'revs', {}).get('PE_sign_name')
            OPINDEX_qa = self.roctop[actpos][cnt].get(
                'revs', {}).get('QA_sign_name')
            OPINDEX_pd = self.roctop[actpos][cnt].get(
                'revs', {}).get('PD_sign_name')
            OPINDEX_date = self.roctop[actpos][cnt].get('revs', {}).get('date')
            OPINDEX_date = OPINDEX_date.replace(".", "/")
            if OPINDEX_IMPL2 is not None:
                impl = OPINDEX_IMPL2
            if OPINDEX_ISS is None:
                OPINDEX_ISS = "see PDF"
            if OPINDEX_qa is None:
                OPINDEX_qa = "see PDF"
            if OPINDEX_pd is None:
                OPINDEX_pd = "see PDF"
            OPINDEX_stageno = self.roctop[actpos][0].get(
                'details', {}).get('stageno')
            OPINDEX_pekpn = self.roctop[actpos][0].get(
                'details', {}).get('pekpn')
            OPINDEX_cstpn = self.roctop[actpos][0].get(
                'details', {}).get('cstpn')
            OPINDEX_copies = self.roctop[actpos][cnt].get(
                'revs', {}).get('copies')
            OPINDEX_rfc = self.roctop[actpos][cnt].get('revs', {}).get('rfc')
            OPINDEX_doc = self.roctop[actpos][cnt].get('revs', {}).get('doc')
            OPINDEX_rcvd = self.roctop[actpos][cnt].get('revs', {}).get('rcvd')
            OPINDEX_pages = self.roctop[actpos][cnt].get(
                'revs', {}).get('pages')
            out = safestr(opno + "_" + rev + ".pdf")
            proj = opno2proj(opno)
            appdirfil = ''
            link2 = "not found"
            if proj != '':
                appdirfil = os.path.abspath(wordy.approvedpath + '\\'
                                            + proj + '\\' + out)
            if os.path.isfile(appdirfil):
                link2 = appdirfil
            return[proj, opno, rev, desc, impl, OPINDEX_ISS, OPINDEX_qa, OPINDEX_pd, OPINDEX_date, OPINDEX_doc, OPINDEX_rfc, OPINDEX_rcvd, OPINDEX_pages, OPINDEX_copies, OPINDEX_stageno, OPINDEX_pekpn, OPINDEX_cstpn, link2]
        else:
            if cnt - 1 >= 1:
                return self.verify_complete(actpos, cnt - 1)
            else:
                return False
            

    def makeXML(self):
        if self.DBOK():
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
            rvs = []
            for  rec in progressbar(self.roctop):
                rvs = []
                savfil = safestr(rec[0]['details']['opno'])
                obstat = safestr(rec[0]['details'].get('obsolete', False))
                rvs.append({'details': {'opno': savfil}})
                rvs.append({'details': {'obsolete': obstat}})
                rvs.append(
                    {'details': {'cstpn': safestr(rec[0]['details']['cstpn'])}})
                rvs.append(
                    {'details': {'pekpn': safestr(rec[0]['details']['pekpn'])}})
                rvs.append(
                    {'details': {'desc': safestr(rec[0]['details']['desc'])}})
                rvs.append(
                    {'details': {'stageno': safestr(rec[0]['details']['stageno'])}})
                rvs.append(
                    {'details': {'stagedesc': safestr(rec[0]['details']['stagedesc'])}})
                found2 = foolproof_finder(rec, 1)
                if found2 > 0:
                    Fu = 1
                    while Fu <= found2:
                        rvs.append({'revs': {'rev': safestr(rec[Fu]['revs']['rev']), 'doc': safestr(rec[Fu]['revs']['doc']), 'rfc': safestr(rec[Fu]['revs']['rfc']), 'pages': safestr(rec[Fu]['revs']['pages']), 'copies': safestr(rec[Fu]['revs']['copies']), 'QA_sign': safestr(
                            rec[Fu]['revs']['QA_sign']), 'PE_sign': safestr(rec[Fu]['revs']['PE_sign']), 'PD_sign': safestr(rec[Fu]['revs']['PD_sign']), 'rcvd': safestr(rec[Fu]['revs']['rcvd']), 'impl': safestr(rec[Fu]['revs']['impl']), 'date': safestr(rec[Fu]['revs']['date'])}})
                        Fu += 1
                xml = dicttoxml.dicttoxml(rvs, attr_type=False)
                xml = xml.replace('<?xml version="1.0" encoding="UTF-8" ?>',
                                  '<?xml version="1.0" encoding="UTF-8" ?><?xml-stylesheet type="text/xsl" href="test2.xsl"?>')

                ROCsave = os.path.abspath(
                    self.ROCpath + '\\' + savfil + ".xml")
                ROCsave1 = os.path.abspath(
                    self.ROCpath + '\\' + "test2.xsl")
                ROCsave2 = os.path.abspath(
                    self.ROCpath + '\\' + savfil + ".html")               
                xmlfile = open(ROCsave, "w")
                xmlfile.write(xml.decode())
                xmlfile.close()
                xslt_doc = etree.parse(ROCsave1)
                xslt_transformer = etree.XSLT(xslt_doc)
                source_doc = etree.parse(ROCsave)
                output_doc = xslt_transformer(source_doc)
                output_doc.write(ROCsave2, pretty_print=True)
                self.ROCxTOCxHTML()

    def ROCxTOCxHTML(self):
        PagesHere = os.listdir(os.path.abspath(self.ROCpath))
        ROCsave3 = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
        with open(ROCsave3, "w") as f:
            f.write(
                "<html><head><link rel='stylesheet' href='listscss.css'></head><body>\n")
            f.write("<h2>Printable Record of Change</h2><ul>")
            last = ''
            prange = 1
            jswitch = 0
            for filename in PagesHere:
                if filename.endswith(".html"):
                    if "TOC.html" not in filename:
                        sfnum, shortfile2, shortfile = self.dig_deep(filename)
                        if last == '':
                            while not (prange <= sfnum <= (prange + 99)):
                                prange = self.add_100(prange)
                                jswitch = 1
                        if not prange <= sfnum <= (prange + 99):
                            while prange < sfnum:
                                prange = self.add_100(prange)
                                if prange > sfnum:
                                    prange = prange - 100
                                    jswitch = 1
                                    break
                        if prange <= sfnum <= (prange + 99):
                            if last != '':
                                f.write("</ul></details>")
                                f.write("</ul></details>")
                            catt = str(str(prange) + " to " + str(prange + 99))
                            f.write(
                                "<details><summary>%s</summary><ul>" % (catt))
                            prange = self.add_100(prange)
                            jswitch = 1
                        if shortfile2 != last:
                            if last != '':
                                if jswitch == 0:
                                    f.write("</ul></details>")
                                elif jswitch == 1:
                                    jswitch = 0
                            f.write("<details><summary>%s</summary><ul>" %
                                    (shortfile2))
                        f.write("<li><a href='%s'>%s</a></li>" %
                                (filename, shortfile))
                        last = shortfile2
            f.write("</ul></details>")
            f.write("</ul></body></html>\n")

    def add_100(self, prange):
        prange = prange + 100
        return prange

    def dig_deep(self, filename):
        res = filename.rfind("PEOI")
        shortfile = filename[res:]
        shortfile2 = shortfile[5: 9]
        sfnum = int(shortfile2)
        return sfnum, shortfile2, shortfile

    def DBpoplist(self):
        if self.DBOK():
            tlist = []
            tlist2 = ['PEOI']
            tlist3=[(0,0)]
            tlist4 =[(0,0)]
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
            n = 1
            for x in self.roctop:
                tlist.append(x[0]['details']['opno'])
                (fx, fx2) = self.find_proj(x[0]['details']['opno'])
                if fx not in tlist2:
                    tlist2.append(fx)
                    tlist3.append((fx, n))
                    try:
                        tlist4.append((int(fx), n))
                    except:
                        tlist4.append((999999999, n))
                    n+=1
            tlist4.sort()        
            tlist5=sorted(tlist4, reverse=True)
            tlist5.pop()
            tlist5.insert(0, (0,0))
            if self.sortorder.GetValue():
                masterlist = tlist4
            else:
                masterlist = tlist5
            list5=[None for _ in range(len(masterlist))]
            for xx in range(len(masterlist)):
                list5[xx]=tlist3[masterlist[xx][1]][0]
            list5[0]="PEOI"
            return (tlist, list5)

    def filter_revs(self, event):
        xx = safestr(self.m_comboBox2qq.GetValue())
        tlist3 = []
        found_you = findme(self.roctop, xx)
        found2 = foolproof_finder(self.roctop[found_you], 1)
        if found2 > 0:
            Fu = 1
            while Fu <= found2:
                tlist3.append(self.roctop[found_you][Fu]['revs']['rev'])
                Fu += 1
        self.tlist3 = tlist3

        self.m_rev_master.Clear()
        self.m_rev_master.AppendItems(tlist3)
        return

    def filter_projs(self, event):
        self.m_rev_master.Clear()
        self.m_rev_master.AppendItems([''])
        tlist = []
        xx = safestr(self.m_comboBox2TL.GetValue())
        for x in self.m_comboBoxDB:
            if x.find(xx) != -1:
                tlist.append(x)
        self.m_comboBox2qq.Clear()
        self.m_comboBox2qq.AppendItems(tlist)
        return

    def Zconvert(self, valIN):
        if valIN == 0:
            return 'awaiting'
        elif valIN == 1:
            return 'approved'
        elif valIN == 2:
            return 'rejected'
        else:
            return 'error'

    def poplist(self, return_early=False):
        x = []
        xx = []
        xxx = []
        zz = time()

        for (dirpath, dirnames, fi) in os.walk(wordy.forappralpath):
            for fp in fi:
                if fp.endswith('.db'):
                    try:
                        killitdead = os.path.abspath(dirpath + "\\" + fp)
                        print(killitdead)
                        if fileaccesscheck(killitdead):
                            os.remove(killitdead)
                    except Exception as e:
                        print('oops ' + str((inspect.stack()[0][2])))
                        print (e.message, e.args)
                        pass
                if not fp.endswith(('.txt', '.db')):
                    x.append(fp)
                    fpx = rreplace(fp, '.pdf', '.txt', 1)
                    fifpx = os.path.abspath(dirpath + "\\" + fpx)
                    xpx = os.path.abspath(dirpath + "\\" + fp)
                    created = os.path.getctime(xpx)
                    try:
                        elapsedraw = zz - created
                        elapsedint = int(
                            ceil((float(elapsedraw) / 3600.00) / 24.00))
                        rejdatax = (str(
                            ' days on system: ' + str(elapsedint)))
                    except Exception as e:
                        print('oops ' + str((inspect.stack()[0][2])))
                        print (e.message, e.args)
                        rejdatax = (str(
                            ' created: ' + str(ctime(created))))
                    if os.path.exists(fifpx):
                        with open(fifpx, "r") as myfile:
                            rejdata = myfile.readlines()
                        xxx.append(rejdata)
                        xrejdata = list(rejdata)
                        xrejdata.append(str(rejdatax))
                        xx.append(xrejdata)
                    else:
                        xxx.append(["", " "])
                        xx.append(rejdatax)

            if fi:
                pass
            if not fi:
                try:
                    if dirpath != wordy.forappralpath:
                        os.rmdir(dirpath)
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    pass
        if self.DBOK():
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
        rec = []
        try:
            pp = 0
            for p in x:
                z1, z2 = getRnP(p)
                z3 = findme(self.roctop, z2)
                found2 = foolproof_finder(self.roctop[z3], 1)
                if found2 > 0:
                    valx = 0
                    valw = -1 # this line is untested
                    for kx in self.roctop[z3]:
                        for kz in kx:
                            if 'revs' in kz:
                                if kx['revs']['rev'] == z1:
                                    valw = valx
                                break
                        valx += 1

                    try:
                        pointless = self.roctop[z3][valw].get('revs')
                    except IndexError:
                        valw = -1

                    z4a = self.Zconvert(
                        self.roctop[z3][valw]['revs']['PE_sign'])
                    z4b = self.Zconvert(
                        self.roctop[z3][valw]['revs']['QA_sign'])
                    z4c = self.Zconvert(
                        self.roctop[z3][valw]['revs']['PD_sign'])

                    entered_date = self.roctop[z3][valw].get(
                        'revs', {}).get('date_entered')

                    implby = self.roctop[z3][valw].get(
                        'revs', {}).get('impl')
                    if len(implby) < 1:
                        implby = 'missing'

                    if return_early:
                        DoC = self.roctop[z3][valw].get(
                            'revs', {}).get('doc')
                        RfC = self.roctop[z3][valw].get(
                            'revs', {}).get('rfc')

                    if entered_date:
                        elapsedraw = zz - float(entered_date)
                        elapsedint = int(
                            ceil((float(elapsedraw) / 3600.00) / 24.00))
                        date_entered2 = str(
                            ' Days on system: ' + str(elapsedint))
                        xxx[pp].append(date_entered2)
                        if return_early:
                            res = {'PEOI': z2, 'rev': z1, 'PE': z4a,
                                   'QA': z4b, 'PD': z4c, 'reject_reason': xxx[pp], 'doc': DoC, 'rfc': RfC, 'implemented': implby}
                        else:
                            res = {'PEOI': z2, 'rev': z1, 'PE': z4a,
                                   'QA': z4b, 'PD': z4c, 'reject_reason': xxx[pp], 'implemented': implby}

                    else:
                        if return_early:
                            res = {'PEOI': z2, 'rev': z1, 'PE': z4a,
                                   'QA': z4b, 'PD': z4c, 'reject_reason': xx[pp], 'doc': DoC, 'rfc': RfC, 'implemented': implby}
                        else:
                            res = {'PEOI': z2, 'rev': z1, 'PE': z4a,
                                   'QA': z4b, 'PD': z4c, 'reject_reason': xx[pp], 'implemented': implby}

                    rec.append(res)
                    pp += 1
            if return_early:
                return rec
            thetime = datetime.datetime.now()
            rec.append({'updatetime': thetime.strftime("%b %d %Y %H:%M:%S")})
            xml = dicttoxml.dicttoxml(rec, attr_type=False)
            xml = xml.replace('<?xml version="1.0" encoding="UTF-8" ?>',
                              '<?xml version="1.0" encoding="UTF-8" ?><?xml-stylesheet type="text/xsl" href="test3.xsl"?>')
            ROCsave = os.path.abspath(self.rootpath + '\\' + "outstanding.xml")
            ROCsave1 = os.path.abspath(self.rootpath + '\\' + "test3.xsl")
            ROCsave2 = os.path.abspath(
                self.rootpath + '\\' + "outstanding.html")
            xmlfile = open(ROCsave, "w")
            xmlfile.write(xml.decode())
            xmlfile.close()
            xslt_doc = etree.parse(ROCsave1)
            xslt_transformer = etree.XSLT(xslt_doc)
            source_doc = etree.parse(ROCsave)
            output_doc = xslt_transformer(source_doc)
            output_doc.write(ROCsave2, pretty_print=True)

        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
            print('error in generating Outstanding data')
        return x

    def reporter(self, event):
        wordy.reportfunc = 1
        x = FinaliseFrame('nothing')
        return

    def find_proj(self, stringin):
        try:
            proj = opno2proj(stringin)
            filenamey = stringin.rfind('PEOI')
            filenamex = len(stringin)
            filename = stringin[filenamey:filenamex]
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            return ('', '')
        return (proj, filename)

    def m_filePicker1OnFileChanged(self, event):
        try:
            if wordy.m_filePicker1 is not None:
                xx = wordy.m_filePicker1
            else:
                xx = safestr(self.m_filePicker1.GetPath())
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            xx = safestr(self.m_filePicker1.GetPath())
        proj = ''
        filename = ''
        if os.path.isfile(xx):
            (proj, filename) = self.find_proj(xx)
        else:
            return
        if proj != '':
            if filename != '':
                obspath = os.path.abspath(self.archive_path + '\\'
                                          + proj)
                self.makepaths(obspath)
            else:
                return
        else:
            return
        obsolete_filepath_file = os.path.abspath(obspath + '\\'
                                                 + filename)
        dlg = wx.MessageDialog(None,
                               'please confirm you wish to make %s obsolete, and move it to %s'
                               % (filename, obspath),
                               'Confirm Archiving', wx.OK | wx.CANCEL)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.pdf_wm_obs(xx, obsolete_filepath_file)
            msg = wx.BusyInfo('Complete - %s has been archived'
                              % filename)
            sleep(0.6)
            try:
                self.m_filePicker1.SetPath('')
                event.Skip()
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                pass
        return

    def pdf_wm_obs(self, approved_filepath_file,
                   obsolete_filepath_file):
        msg = wx.BusyInfo('Working...')
        with open(approved_filepath_file, 'rb') as filehandle_input:

            # read content of the original file

            pdf = PdfFileReader(filehandle_input)
            PDF_data_img = gen_obs()
            # with BytesIO(self.PDF_data) as filehandle_watermark:
            with BytesIO(PDF_data_img) as filehandle_watermark:

                # read content of the watermark

                watermark = PdfFileReader(filehandle_watermark)
                pdf_writer = PdfFileWriter()

                # get first page of the original PDF

                for x in range(0, pdf.getNumPages()):
                    first_page = pdf.getPage(x)

                # get first page of the watermark PDF

                    first_page_watermark = watermark.getPage(0)

                # merge the two pages

                    first_page.mergePage(first_page_watermark)

                # add page

                    pdf_writer.addPage(first_page)
                    x += 1

                    with open(wordy.output_file, 'wb') as \
                            filehandle_output:

                        # write the watermarked file to the new file

                        pdf_writer.write(filehandle_output)

        if os.path.isfile(approved_filepath_file):
            os.remove(approved_filepath_file)
        if os.path.isfile(wordy.output_file):
            if fileaccesscheck(wordy.output_file):
                if not fileaccesscheck(obsolete_filepath_file):
                    os.rename(wordy.output_file, obsolete_filepath_file)
        if os.path.isfile(wordy.output_file):
            if fileaccesscheck(wordy.output_file):
                os.remove(wordy.output_file)
        return

    def on_new_frame(self, event):
        if wordy.frame_number == 0:
            if event == 1:
                wordy.do_save = 1
            else:
                wordy.do_save = 0
            title = 'Record of Change to Operator Instructions'
            frame = OtherFrame(title=title)
            wordy.frame_number = 1

    def openROC(self, event):
        inf = self.m_comboBox2.GetValue()
        if inf != '':
            if inf in self.m_comboBox2ChoicesFULL:
                REV, PEOI = getRnP(inf, True)
                lb = PEOI.find('-')
                if lb == 4:
                    wordy.extractedID = [safestr('PEOI-' + PEOI), REV,
                                         None]
                    self.on_new_frame(0)
                    return

    def approve_frame(self, event, origin):
        if wordy.app_frame_number == 0:
            wordy.origin = origin
            title = 'Approve or Reject the PEOI'
            frame = ApproveFrame(title=title)
            wordy.app_frame_number = 1

    def fin(self, event):
        if wordy.final_frame_number == 0:
            title = 'Check ROC before finalising'
            frame = FinaliseFrame(title=title)
            wordy.final_frame_number = 1
        self.update_the_ROC_list()
        return

    def process_exists(self, process_name, count):
        if count == 0:
            process_name = process_name.upper()
        else:
            process_name = process_name.lower()
        call = ('TASKLIST', '/FI', 'imagename eq %s' % process_name)

        # use buildin check_output right away

        output = SP_check_output(call)

        # check in last line for process name

        last_line = output.strip().split('\r\n')[-1]

        # because Fail message could be translated

        if last_line.startswith('INFO'):
            if count > 0:
                return False
            else:
                self.process_exists(process_name, 1)
                return False
        pn = process_name[0:len(process_name) - 1]

        # return last_line.lower().startswith(process_name.lower())

        if last_line.lower().startswith(pn.lower()):
            return True

    def AskUser(self):
        if self.process_exists('WINWORD*', 0):
            dlg = wx.MessageDialog(None,
                                   'Warning - save and close any Word documents you have open before pressing Okay. If you do not want to do this, press cancel and this program will close.', 'close Word', wx.OK | wx.CANCEL)
            result = dlg.ShowModal()
            dlg.Destroy()
            if result == wx.ID_OK:
                try:
                    os.system('taskkill /f /im  winword* /t')
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    pass
            else:
                self.Destroy()
                self.quit()

    def quit(self):
        try:
            shutil.rmtree(wordy.temp_directory, ignore_errors=True)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        os.sys.exit()

    def in2PDF(self, event):
        wordy.reset_after_reimport = 0
        if os.path.isfile(self.transition_file):
            if self.checkvalid2(self.extracted):
                out_dir = safestr(self.extracted[0] + '_'
                                  + self.extracted[1])
                out = safestr(self.extracted[0] + '_' + self.extracted[1]
                              + '.pdf')
                clout = safestr(
                    self.extracted[0] + '_' + self.extracted[1] + '.txt')
                if not os.path.exists(self.forapprovalpath + '\\'
                                      + out_dir):
                    os.mkdir(self.forapprovalpath + '\\' + out_dir)
                out_file = os.path.abspath(self.forapprovalpath + '\\'
                                           + out_dir + '\\' + out)
                clout_file = os.path.abspath(self.forapprovalpath + '\\'
                                             + out_dir + '\\' + clout)
                proj = opno2proj(out)
                if proj != '':
                    appdirfil = os.path.abspath(wordy.approvedpath + '\\'
                                                + proj + '\\' + out)
                if os.path.isfile(appdirfil):
                    msg = wx.BusyInfo(
                        'This file already exists in approved directory! cancelling import')
                    sleep(4)
                    self.set_blanks()
                    return

                if os.path.isfile(out_file):
                    if (wx.MessageBox(("File exists at awaiting approval stage. Are you sure you want to overwrite?"), ("File Exists!"), parent=self, style=wx.YES_NO | wx.ICON_WARNING,) == wx.NO):
                        msg = wx.BusyInfo('Cancelling...')
                        sleep(1)
                        self.set_blanks()
                        return
                    else:
                        wordy.reset_after_reimport = 1
                        
                if self.lookforobsflag():
                    if (wx.MessageBox(("this PEOI number is marked obsolete, are you sure you want to import?"), ("PEOI OBSOLETE!"), parent=self, style=wx.YES_NO | wx.ICON_WARNING,) == wx.NO):
                        msg = wx.BusyInfo('Cancelling...')
                        sleep(1)
                        self.set_blanks()
                        return               

                msg = wx.BusyInfo('Please wait, %s is being created'
                                  % out)
                self.transition(self.transition_file2, out_file, 17,
                                r"\\NT4\Client_Files\Public\PEOI\notafile.bananananana")
                if os.path.exists(out_file):
                    if self.set_blanks():
                        if os.path.exists(clout_file):
                            try:
                                if fileaccesscheck(clout_file):
                                    os.remove(clout_file)
                            except Exception as e:
                                print('oops ' + str((inspect.stack()[0][2])))
                                print (e.message, e.args)
                                pass
                    self.update_the_ROC_list()
                else:

                    # errror message here

                    return
            else:

                print ("data invalid - cancelling import")
                msg = wx.BusyInfo(
                    'Invalid PEOI / ISSUE / DATE found - please check the Header and Footer in the Word document carefully')
                sleep(4)

                return
        else:

            # error message here

            return
        wordy.extractedID = self.extracted
        self.on_new_frame(1)
        self.update_the_ROC_list()
        return

    def set_blanks(self):
        self.m_textCtrl4.SetValue('blank')
        self.m_textCtrl42.SetValue('blank')
        self.m_textCtrl43.SetValue('blank')
        if os.path.exists(self.transition_file):
            try:
                if fileaccesscheck(self.transition_file):
                    os.remove(self.transition_file)
                if fileaccesscheck(self.transition_file2):
                    os.remove(self.transition_file2)
                return True
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)
                pass
                return False
        return
    
    def lookforobsflag(self):
        kill = "z"
        xcount = 0
        for x in self.roctop:
            if x[0]['details']['opno'] == self.PEOI:
                kill = xcount
                break
            xcount += 1
        if kill == "z":
            return False
        if kill != "z":
            if self.roctop[kill][0]['details'].get("obsolete"):
                return True
            else:
                return False

    def update_the_ROC_list(self, event=None):
        self.m_comboBox2ChoicesFULL = self.poplist()
        self.m_comboBox2.Clear()
        self.m_comboBox2.AppendItems(self.m_comboBox2ChoicesFULL)
        (self.m_comboBoxDB, self.m_comboBoxDBTL) = self.DBpoplist()
        self.m_comboBox2qq.Clear()
        self.m_comboBox2qq.AppendItems(self.m_comboBoxDB)
        self.m_comboBox2TL.Clear()
        self.m_comboBox2TL.AppendItems(self.m_comboBoxDBTL)
        self.m_rev_master.Clear()
        self.m_rev_master.AppendItems([''])
        wordy.picklist = self.poplist()
        return

    def import_original(self, event):
        self.AskUser()
        in_file = self.get_path('*.doc*')
        if in_file is None:
            return
        if not os.path.isfile(in_file):
            return
        msg = wx.BusyInfo('Please wait, %s conversion in progress...'
                          % in_file)
        self.transition(in_file, self.transition_file,
                        16, self.transition_file2)
        text = safestr(DOC2_proc(self.transition_file))
        self.PEOI = self.digger(text, 'PEOI', 0)
        xx = (text.split("\n"))
        indices = [i for i, s in enumerate(
            xx) if 'For Approval Signatures' in s]
        twelve_step = self.digger(text, 'DATE', 12)
        twelve_step_iss = self.digger(text, 'ISSUE', 12)
        if len(indices) > 0:
            ind = indices[0]
            self.ISSUE = safestr(xx[ind - 4])
            self.DATE = safestr(xx[ind - 2])
        else:
            self.ISSUE = twelve_step_iss
            self.DATE = twelve_step
        valid = self.roundup()
        if len(safestr(self.ISSUE)) > 2 or len(safestr(self.ISSUE)) < 1:
            self.ISSUE = twelve_step_iss
            self.DATE = twelve_step
            valid = self.roundup()
        if len(safestr(self.ISSUE)) > 2 or len(safestr(self.ISSUE)) < 1:
            self.ISSUE = self.digger(text, 'ISSUE', 10)
            valid = self.roundup()
        if len(safestr(self.ISSUE)) > 2 or len(safestr(self.ISSUE)) < 1:
            if self.ISSUE.upper().find('DRAFT') != -1 or self.ISSUE.upper().find('PAGE') != -1:
                self.ISSUE = 'INVALID'
            valid = self.roundup()
        if self.DATE.find('Q.A.') != -1:
            self.DATE = self.digger(text, 'DATE', 10)
            valid = self.roundup()
        if valid:
            self.m_textCtrl4.SetValue(self.PEOI)
            self.m_textCtrl42.SetValue(self.ISSUE)
            self.m_textCtrl43.SetValue(self.DATE.strip())
        else:
            self.ISSUE = twelve_step_iss
            self.DATE = twelve_step
            valid = self.roundup()
            if valid:
                self.m_textCtrl4.SetValue(self.PEOI)
                self.m_textCtrl42.SetValue(self.ISSUE)
                self.m_textCtrl43.SetValue(self.DATE.strip())
            else:
                self.m_textCtrl4.SetValue('error')
                self.m_textCtrl42.SetValue('error')
                self.m_textCtrl43.SetValue('error')
                
                     
        # other data could be extracted from the word doc and stored for
        # later retieval (ie when populating the Record of Change for the first time)
        # the 3 lines below are unfiltered examples
        desc = self.filter_me(safestr(self.digger(text, 'DESCRIPTION', 0)))
        cpn = self.filter_me(safestr(self.digger(text, 'CUSTOMER PART NUMBER', 0)))
        snsd = self.filter_me(safestr(self.digger(text, 'STAGE NO', 0)))
        pn = self.filter_me(safestr(self.digger(text, 'A-', 0)))
        if pn == "4" or pn == 4:
           pn = self.filter_me(safestr(self.digger(text, 'ASS-', 0)))
        if pn == "4" or pn == 4:
           pn = self.filter_me(safestr(self.digger(text, 'SUB-', 0)))
        
        sn = snsd
        sd = snsd
        try:
            snsplit = next(i for i,j in list(enumerate(snsd,1))[::-1] if j.isdigit())          
        except StopIteration:
            try:
                snsplit = next(i for i,j in list(enumerate(snsd,1))[::-1] if j in [",", "-"])-1
            except StopIteration:
                snsplit = 0
        if snsplit !=0: 
            sn = sn[:snsplit].strip()
            sd = sd[snsplit:].strip()
            for charz in ",-":
                sd = sd.replace(charz,"").strip()                  
        try:
            wordy.headerstuff = [self.PEOI, desc, cpn, sn, sd, pn]
        except NameError:
            wordy.headerstuff = False
        return
    
    def filter_me(self, string):
        index = string.find(":")
        if index > 0:
             return (string[index:]).replace(":","").strip()
        else:
            return string


    def roundup(self):
        if self.checkvalid([self.PEOI, self.ISSUE, self.DATE]):
            self.extracted = self.cleanup([self.PEOI, self.ISSUE,
                                           self.DATE])
            self.PEOI = safestr(self.extracted[0])
            self.ISSUE = safestr(self.extracted[1])
            self.DATE = safestr(self.extracted[2])
            return True
        else:
            return False

    def textupdated(self, event):
        if self.m_textCtrl43.GetValue() == 'blank':
            return
        if self.m_textCtrl42.GetValue() == 'blank':
            return
        if self.m_textCtrl4.GetValue() == 'blank':
            return
        self.extracted = self.cleanup([self.m_textCtrl4.GetValue(),
                                       self.m_textCtrl42.GetValue(),
                                       self.m_textCtrl43.GetValue()])
        wordy.final_filename = str(
            safestr(self.extracted[0]) + '_' + safestr(self.extracted[1]))
        self.m_staticText241.SetLabel(wordy.final_filename + '.pdf')
        return

    def cleanup(self, varz):
        PEOI = varz[0]
        ISSUE = varz[1]
        DATE = varz[2]
        badthings = r'[:"*?<>|]+'
        goodthings = r"_-"
        PEOI2 = '_'.join(PEOI.split())
        ISSUE2 = '_'.join(ISSUE.split())
        safePEOI = re_sub(r'[^\w' + goodthings + ']', '',
                          safestr(PEOI2))
        safePEOI = ''.join(safePEOI.split())
        safeISSUE = re_sub(r'[^\w' + goodthings + ']', '',
                           safestr(ISSUE2))
        safeISSUE = ''.join(safeISSUE.split())
        safeDATE = re_sub(badthings, '', DATE)
        safeDATE = ''.join(safeDATE.split())
        return [safePEOI.upper(), safeISSUE.upper(), safeDATE.upper()]

    def checkvalid(self, valz):
        for chk in valz:
            if len(chk) > 0:
                pass
            else:
                return False
        return True

    def checkvalid2(self, valz):
        for chk in valz:
            if chk == "INVALID":
                return False
            else:
                pass
        return True

    def digger(
        self,
        txt,
        val,
        offset,
    ):
        return txt[txt.find(val):].split('\n')[offset]

    def transition(
        self,
        file_in,
        file_out,
        format,
        file_out2,
    ):

        # wdFormatPDF = 17 wdFormatDOCX = 16
        word = CreateObject('Word.Application')
        sleep(3)
        word.Visible = False
        if os.path.isfile(file_in):
            doc = word.Documents.Open(file_in)
            if not os.path.isfile(file_out):
                pass
            else:
                try:
                    if fileaccesscheck(file_out):
                        os.remove(file_out)
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    return
            if not os.path.isfile(file_out2):
                pass
            else:
                try:
                    if fileaccesscheck(file_out2):
                        os.remove(file_out2)
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
                    return
        else:
            return
        if not file_out.endswith('.pdf'):
            doc.SaveAs(file_out, FileFormat=format)
            shutil.copy(file_in, file_out2)
        else:
            doc.ExportAsFixedFormat(OutputFileName=file_out,
                                    ExportFormat=17,  # 17 = PDF output, 18=XPS output
                                    OpenAfterExport=False,
                                    # 0=Print (higher res), 1=Screen (lower res)
                                    OptimizeFor=0,
                                    # 0=No bookmarks, 1=Heading bookmarks only, 2=bookmarks match word bookmarks
                                    CreateBookmarks=0,
                                    DocStructureTags=False,
                                    IncludeDocProps=False,
                                    )
        if file_out.endswith('.pdf'):
            wordy.file_out = file_out
        doc.Close()
        word.Quit()
        # ammendments start here - adding metadata
        if file_out.endswith('.pdf'):
            fin = open(file_out, 'rb')
            reader = PdfFileReader(fin)
            writer = PdfFileWriter()
            #only add a page if it isn't blank (determined by containing some text) START
            NumPages = reader.getNumPages()
            for i in range(0, NumPages):
                PageObj = reader.getPage(i)
                Text = PageObj.extractText().split()
                if len(Text)>1:
                    writer.addPage(PageObj)
            #only add a page if it isn't blank (determined by containing some text) END
            #writer.appendPagesFromReader(reader) # don't need this line - this added all pages regardless
            metadata = reader.getDocumentInfo()
            writer.addMetadata(metadata)
            writer.addMetadata({
                '/Author': 'Production Engineering',
                '/Title': safestr(wordy.final_filename)
            })
            fout = open(file_out, 'ab')
            writer.write(fout)
            fin.close()
            fout.close()
            # ammendments end here
        return

    def get_path(self, wildcard):
        style = wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        dialog = wx.FileDialog(self, 'Open', wildcard=wildcard,
                               style=style)
        if dialog.ShowModal() == wx.ID_OK:
            path = dialog.GetPath()
        else:
            path = None
        dialog.Destroy()
        return path

    def makepaths(self, x):
        if not os.path.exists(x):
            try:
                os.makedirs(x)
            except OSError, exc:

                # Guard against race condition

                pass

    def xferttech(self, Xto, Xfrom):
        try:
            distutils.dir_util.remove_tree(Xto)
        except Exception as e:
            print('oops')
            print (e.message, e.args)
        try:
            distutils.dir_util.copy_tree(Xfrom, Xto)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        return

    def __del__(self):
        try:
            self.poplist()
            # self.makeXML()
            if fileaccesscheck(self.transition_file):
                os.remove(self.transition_file)
            if fileaccesscheck(self.transition_file2):
                os.remove(self.transition_file2)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)


class GodMode_Frame(wx.Frame):

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title=title,
            pos=wx.DefaultPosition,
            size=wx.Size(224, 350),
            style=wx.DEFAULT_FRAME_STYLE & ~(
                wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
        )

        self.godmode_string = r'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'
        self.godmode_data = b64decode(self.godmode_string)
        self.stream = BytesIO(bytearray(self.godmode_data))
        self.GMimage = wx.Image(self.stream, wx.BITMAP_TYPE_ANY)
        self.GODBMP = self.GMimage.ConvertToBitmap()
        self.gm_bitmap = wx.StaticBitmap(self,
                                         wx.ID_ANY,
                                         wx.Bitmap(self.GODBMP),
                                         wx.DefaultPosition,
                                         wx.DefaultSize,
                                         0,
                                         )
        self.SetBackgroundColour(wx.BLACK)
        self.SetForegroundColour((95, 90, 90))
        self.warning = wx.StaticText(
            self,
            wx.ID_ANY,
            u'Warning: God Mode Enabled:  ',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.warning2 = wx.StaticText(
            self,
            wx.ID_ANY,
            u'You have full-access to all department functions, can bulk-add R.O.C. revs, sign-off from the R.O.C. window for all departments etc',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.warning2.Wrap(216)
        bSizer3 = wx.BoxSizer(wx.VERTICAL)
        bSizer3.Add(self.warning, 1, wx.EXPAND, 5)
        bSizer3.Add(self.gm_bitmap, 1,  wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer3.Add(self.warning2, 1, wx.EXPAND, 5)
        self.SetSizer(bSizer3)
        self.Layout()
        self.arrayfun = [["ayTNMo3cmTMM9", "yaTNM3ocmMMM9"], ["RVvbTjFRB", "RVvbjTFRB"], ["IQMJKJC7", "IQJMKJ7C"], ["EP9BxOnecuPrn", "EP9BxOencurPn"], ["EXgYwd1WAd1A", "EXgwYd1AWd1A"], ["dmnXHYIJroie", "dnmXHYIJorie"], ["TwujIyqgE1GupkpwB", "TwujIygqE1GukppwB"], [
            "BNMxLiNyvawSywYUrHITQO", "BMNxLiNvyawSywUYrHITQO"], ["FdKJDamMtim", "FeKKDacMtim"], ["aeC4V5SGLxQ2mvvFDtsEU", "aeF4V5SGLxQ7mvvFDhsEU"], ["wappgvxei4kN8L6DyZnPq", "wappgaxei3kN8L3DyZnPq"], ["z", "z"], ["x", "x"], ["q", "q"]]
        self.Show()

        myThread = Thread(target=self.myTimer)
        myThread.start()

    def myTimer(self):
        self.countdown(10)
        return

    def countdown(self, num):
        try:
            self.warning.SetLabel(
                'Warning: God Mode Enabled: ' + safestr(num))
            num -= 1
            sleep(0.1)
            self.distort(num)
            sleep(0.1)
            self.distort(num)
            sleep(0.1)
            self.distort(num)
            sleep(0.1)
            self.distort(num)
            sleep(0.6)
            if num >= 1:
                self.countdown(num)
            else:
                try:
                    self.Close()
                except Exception as e:
                    print('oops ' + str((inspect.stack()[0][2])))
                    print (e.message, e.args)
        except RuntimeError:
                print("God Mode warning window closed")
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            try:
                self.Close()
            except Exception as e:
                print('oops ' + str((inspect.stack()[0][2])))
                print (e.message, e.args)

        return

    def distort(self, num):
        if num < 4:
            if len(self.arrayfun) > 0:
                newgm = self.godmode_string.replace(
                    self.arrayfun[-1][0], self.arrayfun[-1][1])
                del self.arrayfun[-1]
                self.godmode_data = b64decode(newgm)
                self.stream = BytesIO(bytearray(self.godmode_data))
                self.GMimage = wx.Image(self.stream, wx.BITMAP_TYPE_ANY)
                self.GODBMP = self.GMimage.ConvertToBitmap()
                self.gm_bitmap.SetBitmap(wx.Bitmap(self.GODBMP))
            return


if __name__ == '__main__':
    app = wx.App()
    wordy(None, title='Operator Instructions Change Control')
    print ('GUI booting now')
    print ('you can minimise')
    print ("but don't close")
    print ("this window")

    app.MainLoop()
