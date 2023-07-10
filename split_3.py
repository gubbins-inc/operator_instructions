# split 3

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

