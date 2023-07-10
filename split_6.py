# split 6

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

