# split 8 
        
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

