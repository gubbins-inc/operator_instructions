# split 4

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

