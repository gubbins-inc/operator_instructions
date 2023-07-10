# split 7

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
 
