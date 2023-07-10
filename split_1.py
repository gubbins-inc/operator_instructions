# split 1

class ROC_frame(wx.Frame):

    """
    Class used for creating frames other than the main one
    """

    def __init__(self, title, parent=None):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title='READ-only R.O.C: ' + title,
            pos=wx.DefaultPosition,
            size=wx.Size(1166, 570),
            style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL,
        )
        self.SetIcon(OICC_LOGO.GetIcon())
        self.title = title
        self.panel = wx.Panel(self)
        self.theDB = wordy.DBpath

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        self.roctop = []

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

        self.deftxt = []
        for n in range(0, 7):
            self.deftxt.append(wx.TextCtrl(
                self.panel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0,))

        self.m_textCtrl_PEOI = self.deftxt[0]
        self.m_textCtrl_PEOI.SetEditable(False)
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

        self.m_textCtrl_cust_PN = self.deftxt[1]
        self.m_textCtrl_cust_PN.SetEditable(False)
        self.m_textCtrl_cust_PN.SetBackgroundColour((195, 195, 195))
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

        self.m_textCtrl_pekPN = self.deftxt[2]
        self.m_textCtrl_pekPN.SetEditable(False)
        self.m_textCtrl_pekPN.SetBackgroundColour((195, 195, 195))
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

        self.m_textCtrl_prod_desc = self.deftxt[3]
        self.m_textCtrl_prod_desc.SetEditable(False)
        self.m_textCtrl_prod_desc.SetBackgroundColour((195, 195, 195))
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

        self.m_textCtrl_stages = self.deftxt[4]
        self.m_textCtrl_stages.SetEditable(False)
        self.m_textCtrl_stages.SetBackgroundColour((195, 195, 195))
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

        self.m_textCtrl_proc = self.deftxt[5]
        self.m_textCtrl_proc.SetEditable(False)
        self.m_textCtrl_proc.SetBackgroundColour((195, 195, 195))
        bSizer5.Add(self.m_textCtrl_proc, 1, wx.ALL, 5)

        self.obsstat = wx.StaticText(
            self.panel,
            wx.ID_ANY,
            u'Obsolete?',
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.obsstat.Wrap(-1)
        bSizer5.Add(self.obsstat, 0, wx.ALL, 5)

        self.obs = self.deftxt[6]
        self.obs.SetEditable(False)
        self.obs.SetBackgroundColour((195, 195, 195))
        bSizer5.Add(self.obs, 0, wx.ALL, 5)

        bSizer1.Add(bSizer5, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline3 = wx.StaticLine(self.panel, wx.ID_ANY,
                                           wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer6.Add(self.m_staticline3, 1, wx.ALL | wx.EXPAND, 5)

        bSizer1.Add(bSizer6, 0, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        bSizer1.Add(bSizer7, 0, wx.EXPAND, 5)

        bSizer61 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticline31 = wx.StaticLine(self.panel, wx.ID_ANY,
                                            wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer61.Add(self.m_staticline31, 0, wx.ALL | wx.EXPAND, 5)

        bSizer52 = wx.BoxSizer(wx.VERTICAL)
        bSizer52a = wx.BoxSizer(wx.HORIZONTAL)

        self.roc_prev = WX_Grid(self.panel, wx.ID_ANY,
                                wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid

        rowz = self.rowcalc()
        self.found_you = rowz[1]
        self.rows = rowz[0]
        self.roc_prev.CreateGrid(self.rows, 10)
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
        self.roc_prev.SetColSize(8, 120)
        self.roc_prev.SetColSize(9, 120)
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
        self.roc_prev.SetColLabelValue(8, u'FILE LOCATION')
        self.roc_prev.SetColLabelValue(9, u'link')
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
        self.roc_prev.GetTargetWindow().SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.roc_prev.GetGridWindow().Bind(wx.EVT_MOTION,
                                           self.OnMouseMotion)

        bSizer52.Add(self.roc_prev, 1, wx.ALL | wx.EXPAND, 5)
        bSizer61.Add(bSizer52, 1, wx.EXPAND
                     | wx.ALIGN_CENTER_HORIZONTAL, 5)
        bSizer1.Add(bSizer61, 1, wx.EXPAND
                    | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.roc_prev.HideCol(9)

        self.panel.SetSizer(bSizer1)
        self.panel.Layout()
        self.roc_prev.Bind(WX_EVT_GRID_SELECT_CELL, self.onSingleSelect)

        self.panel.Centre(wx.BOTH)

        # should I move this in to the def?

        self.load_data()
        self.panel.Show()
        self.Show()

    def OnMouseMotion(self, event):
        self.roc_prev.GetTargetWindow().SetCursor(wx.Cursor(wx.CURSOR_HAND))
        event.Skip()

    def onSingleSelect(self, event):
        self.currentlySelectedCell = (event.GetRow(), 9)
        path_to_f = \
            self.roc_prev.GetCellValue(self.currentlySelectedCell)
        if os.path.isfile(path_to_f):
            SP_Popen(path_to_f, shell=True)
        event.Skip()

    def rowcalc(self):
        if not os.path.isfile(self.theDB):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return
        with open(self.theDB, 'rb') as f:
            self.roctop = PKL_load(f)
        found_you = -1
        found2 = -1
        found_you = findme(self.roctop, self.title)
        if found_you >= 0:
            if len(self.roctop[found_you]) > 0:
                found2 = foolproof_finder(self.roctop[found_you], 1)
        return [found2, found_you]

    def load_data(self):
        found_you = self.found_you
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
            self.m_textCtrl_PEOI.SetLabel(safestr(self.roctop[found_you][0]['details'
                                                                            ]['opno'], True))
            self.obs.SetLabel(safestr(self.roctop[found_you][0]['details'].get('obsolete', "False"), True))

            if len(self.roctop[found_you]) > 0:
                acounter = self.rows
                xr = 0
                msg = wx.BusyInfo('Please wait, this may take a while (depends on the number of issue levels released).')
                while acounter >= 1:
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
                        pdf_name = self.title + '_' \
                            + safestr(self.roctop[found_you][acounter]['revs'
                                                                       ]['rev']) + '.pdf'
                        try:
                            (res1, res2) = search_file(wordy.rootpath,
                                                       pdf_name)
                        except TypeError:
                            (res1, res2) = ('missing', '')
                        self.roc_prev.SetCellValue(xr, 9, safestr(res2))
                        self.roc_prev.SetCellValue(xr, 8, safestr(res1))

                    xr += 1
                    acounter -= 1

    def __del__(self):
        wordy.ROROC_frame_number = 0

