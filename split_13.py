# split 13

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

