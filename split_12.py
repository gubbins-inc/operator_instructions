# split 12

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

