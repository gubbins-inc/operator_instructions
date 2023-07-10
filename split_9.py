# split 9

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

