# split 10

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

