# split 5

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

