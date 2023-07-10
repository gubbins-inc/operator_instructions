# split 2

class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):

    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, wx.ID_ANY, style=wx.LC_REPORT |
                             wx.SUNKEN_BORDER)
        CheckListCtrlMixin.__init__(self)
        ListCtrlAutoWidthMixin.__init__(self)


class RcVr(wx.Dialog):

    def __init__(self, *args, **kw):
        super(RcVr, self).__init__(*args, **kw)
        self.SetIcon(OICC_LOGO_5.GetIcon())
        self.isopen = False
        self.SetSize((1000, -1))
        whichwindow = (args[1])
        if whichwindow == 0:
            self.finalised_results = [g for g in wordy.for_rcv_list]
        elif whichwindow == 1:
            self.finalised_results = [g for g in wordy.group_print_list]
        else:
            self.Close()
            self.Destroy()
            return

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        leftPanel = wx.Panel(panel)
        rightPanel = wx.Panel(panel, size=(570, -1))

        self.list = CheckListCtrl(rightPanel)
        self.list.InsertColumn(0, 'PEOI number', width=130)
        self.list.InsertColumn(1, 'Rev', width=50)
        self.list.InsertColumn(2, 'Description', width=260)
        self.list.InsertColumn(3, 'Details of Change', width=260)
        self.list.InsertColumn(4, 'PDF', width=4)
        self.list.InsertColumn(5, 'receiver', width=130)

        idx = 0
        self.receiverlist = []
        for i in self.finalised_results:
            self.receiverlist.append(i[5])
            index = self.list.InsertItem(idx, i[0])
            self.list.SetItem(index, 1, i[1])
            self.list.SetItem(index, 2, i[2])
            self.list.SetItem(index, 3, i[3])
            self.list.SetItem(index, 4, i[4])
            self.list.SetItem(index, 5, i[5])
            idx += 1
        self.receiverlist = list(set(self.receiverlist))

        vbox2 = wx.BoxSizer(wx.VERTICAL)

        config = configparser.ConfigParser()
        config.read('settings.ini')
        try:
            self.me = (safestr(config['user_profile']['me']))
        except KeyError:
            self.me = "UNSET"
        signedbyme = " Files will be marked as \n signed by {}".format(self.me)
        selBtn = wx.Button(leftPanel, label='Select All')
        desBtn = wx.Button(leftPanel, label='Deselect All')
        selByBtn = wx.Button(leftPanel, label='Select By')
        self.recvrscombos = wx.ComboBox(
            leftPanel,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
        )
        for m in self.receiverlist:
            self.recvrscombos.Append(m)
        if whichwindow == 0:
            appBtn = wx.Button(leftPanel, label='Mark Received')
            note = wx.StaticText(leftPanel, label=signedbyme)
            self.Bind(wx.EVT_BUTTON, self.GroupSignFor, id=appBtn.GetId())
        elif whichwindow == 1:
            appBtn = wx.Button(leftPanel, label='Group Open')
            note = wx.StaticText(leftPanel, label="")
            self.Bind(wx.EVT_BUTTON, self.GroupOpen, id=appBtn.GetId())
            appBtn2 = wx.Button(leftPanel, label='txt')
            self.Bind(wx.EVT_BUTTON, self.Checkup, id=appBtn2.GetId())

        self.Bind(wx.EVT_BUTTON, self.OnSelectAll, id=selBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnDeselectAll, id=desBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnSelBy, id=selByBtn.GetId())

        vbox2.Add(selBtn, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(self.recvrscombos, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(selByBtn, 3, wx.EXPAND | wx.BOTTOM, 1)
        vbox2.Add(desBtn, 3, wx.EXPAND, 1)
        # Create a horizontal box sizer
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add the two buttons to the horizontal sizer
        btn_sizer.Add(appBtn, 1, wx.EXPAND, 0)  # Proportion is set to 1, so this button will be resized
        if whichwindow == 1:
            btn_sizer.Add(appBtn2, 0, wx.EXPAND, 0)  # Proportion is set to 0, so this button will not be resized and will take up only the space it needs

        # Add the horizontal sizer to the main vertical sizer (vbox2)
        vbox2.Add(btn_sizer, 3, wx.EXPAND, 1)

        vbox2.Add(note, 3, wx.EXPAND, 5)

        leftPanel.SetSizer(vbox2)

        vbox.Add(self.list, 4, wx.EXPAND | wx.TOP, 3)
        vbox.Add((-1, 10))

        rightPanel.SetSizer(vbox)

        hbox.Add(leftPanel, 0, wx.EXPAND | wx.RIGHT, 5)
        hbox.Add(rightPanel, 1, wx.EXPAND)
        hbox.Add((3, -1))

        panel.SetSizer(hbox)
        if whichwindow == 0:
            self.SetTitle('Mark as Received')
        elif whichwindow == 1:
            self.SetTitle('Group Open for Print')
        self.Centre()

    def OnSelectAll(self, event):
        self.OnDeselectAll(None)
        num = self.list.GetItemCount()
        for i in range(num):
            self.list.CheckItem(i)

    def OnSelBy(self, event):
        self.OnDeselectAll(None)
        if len(self.recvrscombos.GetValue()) > 0:
            num = self.list.GetItemCount()
            for i in range(num):
                if self.recvrscombos.GetValue() == self.list.GetItemText(i, 5):
                    self.list.CheckItem(i)

    def OnDeselectAll(self, event):

        num = self.list.GetItemCount()
        for i in range(num):
            self.list.CheckItem(i, False)

    def GroupSignFor(self, event):
        self.theDB = wordy.DBpath
        (DBOK, self.roctop) = loadUP(self.theDB)
        if not DBOK:
            print("DB error")
            return
        num = self.list.GetItemCount()
        for i in range(num):
            if self.list.IsChecked(i):
                self.add_to_DB(self.list.GetItemText(i, 0),
                               self.list.GetItemText(i, 1))
                msg = wx.BusyInfo(
                    'marking %s rev %s as received by %s' % (self.list.GetItemText(i, 0), self.list.GetItemText(i, 1), self.me))
                sleep(0.20)
        with open(self.theDB, 'wb') as f:
            PKL_dump(self.roctop, f, indent=2)
        print(datetime.datetime.now())
        msg = wx.BusyInfo('Complete!')
        sleep(0.4)
        self.Close()

    def add_to_DB(self, toof, threef):
        kill = "z"
        killer = "z"
        xcount = 0
        for x in self.roctop:
            if x[0]['details']['opno'] == toof:
                kill = xcount
                break
            xcount += 1
        if kill == "z":
            print("PEOI number %s not found." % (toof))
            return
        if kill != "z":
            c = 0
            for z in self.roctop[kill]:
                try:
                    if z['revs']['rev'] == threef:
                        killer = c
                except KeyError:
                    pass
                c += 1
            if killer == "z":
                print ('Revision not there')
                return
            elif killer != "z":
                self.roctop[kill][killer]['revs']['date'] = safestr(
                    datetime.datetime.today().strftime('%d-%m-%Y'))
                self.roctop[kill][killer]['revs']['rcvd'] = self.me

    def shorten(self, string2short, len2short):
        if len(string2short) > len2short:
            return str(string2short[0:len2short-1] + "...")
        else:
            return string2short

    def Checkup(self, event):
        self.GroupOpen(None, called_from_checkup=True)
        self.Close()
        return

    def GroupOpen(self, event, called_from_checkup=False):
        num = self.list.GetItemCount()
        total = 0
        for i in range(num):
            if self.list.IsChecked(i):
                total += 1
        if total == 0:
            return
        csvdetails = ["PEOI NUMBER", "REV", "DESCRIPTION", "DETAILS OF CHANGE"]
        csvrows = []
        for i in range(num):
            if self.list.IsChecked(i):
                templist = []
                x = self.list.GetItemText(i, 4)
                self.manager = self.list.GetItemText(i, 5)
                x1, x2, x3, x4 = self.list.GetItemText(i, 0).strip(), self.list.GetItemText(i, 1).strip(), self.list.GetItemText(i, 2).strip(), self.list.GetItemText(i, 3).strip()
                templist.append(self.shorten(x1, 25))
                templist.append(self.shorten(x2, 25))
                templist.append(self.shorten(x3, 45))
                templist.append(self.shorten(x4, 95))
                csvrows.append(templist)
                fpath = x[:len(x) - 4]
                out_file = os.path.abspath(wordy.forappralpath + '\\'
                                           + fpath + '\\' + x)
                if not called_from_checkup:
                    if os.path.isfile(out_file):
                        SP_Popen([out_file], shell=True)
        col1 = []
        col2 = []
        col3 = []
        col4 = []
        longs = []
        for rows in csvrows:
            col1.append(rows[0])
            col2.append(rows[1])
            col3.append(rows[2])
            col4.append(rows[3])
        cols = [col1, col2, col3, col4]
        for col in cols:
            try:
                longest = max(col, key=len)
                longs.append(len(longest))
            except ValueError:
                longs.append(0)
        for rows in csvrows:
            rows[0] = rows[0].ljust(16)
            rows[1] = rows[1].ljust(8)
            rows[2] = rows[2].ljust(longs[2]+4)
            rows[3] = rows[3].ljust(longs[3]+4)
        csvdetails[0] = csvdetails[0].ljust(16)
        csvdetails[1] = csvdetails[1].ljust(8)
        csvdetails[2] = csvdetails[2].ljust(longs[2]+4)
        csvdetails[3] = csvdetails[3].ljust(longs[3]+4)
        try:
            tf1 = os.path.abspath(wordy.temp_directory + "\\" + 'justprinted.txt')
            tf2 = os.path.abspath(wordy.temp_directory + "\\" + 'justprinted.pdf')
            with open(tf1, 'w') as f:
                write = csv.writer(f, delimiter='\t', doublequote=False, escapechar='\\', quoting=csv.QUOTE_NONE)
                write.writerow(["            OPERATOR INSTRUCTIONS TOP SHEET.     DATE: ", safestr(datetime.datetime.today().strftime('%d-%m-%Y'))])
                write.writerow([self.manager, "- you have been issued with the following (new or updated) laminated instructions:"])
                write.writerow(csvdetails)
                write.writerows(csvrows)
                write.writerow(["Please mark these as received in the OICC software (Finalise > Open Receiver Window)."])
            if called_from_checkup:
                if os.path.isfile(tf1):
                    SP_Popen([tf1], shell=True)
                return
            createpdf = FPDF()
            createpdf.add_page(orientation='L')
            createpdf.set_font("Courier", size=10)
            text_file = open(tf1, "r")
            for text in text_file:
                createpdf.cell(220, 12, txt=text, ln=1, align='L')
            createpdf.output(tf2)
            if os.path.isfile(tf2):
                SP_Popen([tf2], shell=True)
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            print("temp file / PDF gen error")
        self.Close()

