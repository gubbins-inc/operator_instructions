# split 11

    def loggindb(self, event=None):
        if 1 == 1:
            # ask if they are opening to email report, or just to look

            del_or_no = wx.MessageDialog(None,
                                         ('Do you just wish to view the list of recently updated op-ins or distribute the list?') + '\n' +
                                         ('If you click YES, you must copy and paste the text file in to an email and distribute it.') + '\n' +
                                         ('CLick YES if you are going to do this or NO if you just want to view the report.'),
                                         ('Just view report = NO, Distribute via email = YES'),
                                         wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION).ShowModal()
            if del_or_no != wx.ID_YES:
                do_delete = False
            elif del_or_no == wx.ID_YES:
                do_delete = True
            connection_obj = sqlite3.connect(wordy.loggingDB)
            # wordy.loggingrecent
            # cursor object
            # get everything from DB and write it to text file
            outlist = []
            # this gets the DB stuff
            cursor_obj = connection_obj.cursor()
            # connection object - updated_instructions is the 'table'
            cursor_obj.execute("SELECT * FROM updated_instructions")
            # get everything
            x = (cursor_obj.fetchall())
            # get it all in one simple list
            if len(x) > 0:
                for p in x:
                    for j in p:
                        outlist.append(j)
                    outlist.append("")
                    outlist.append("---")
                    outlist.append("")
                # write to text file
                with open(wordy.loggingrecent, 'w') as f:
                    for item in outlist:
                        f.write("%s\n" % item)

                if do_delete:
                    # delete data
                    cursor_obj.execute("DELETE FROM updated_instructions")
                    connection_obj.commit()
                    # Close the connection
                    connection_obj.close()

                    wait = wx.BusyInfo(
                        "Copy & paste text file contents in to an email to report recently updated instructions")
                    sleep(1.3)

                    if os.path.isfile(wordy.email):
                        SP_Popen([wordy.email], shell=True)
                    else:
                        print(wordy.email)
                else:
                    wait = wx.BusyInfo(
                        "These are all instructions updated since the last report was distributed" + '\n' +
                        "You are just viewing this file and have not altered the contents of the databse.")
                    sleep(1.3)
            else:
                wait = wx.BusyInfo(
                    "No New updates were found since this was last run. Now opening the previous report" + '\n' +
                    "... the database indicates these have already been distributed in an email")
                sleep(1.3)
                if do_delete:
                    if os.path.isfile(wordy.email):
                        SP_Popen([wordy.email], shell=True)
                    else:
                        print(wordy.email)

            if os.path.isfile(wordy.loggingrecent):
                SP_Popen([wordy.loggingrecent], shell=True)
            else:
                pass
            del wait
        else:
            pass

    def helpfiles(self, event):
        evt_id = event.GetId()
        options = ['nothing', r"C:\Program Files (x86)\OICC\help\OICC_UM_PE.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_QA.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_PD.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_rectify_rejection.pdf",
                   r"C:\Program Files (x86)\OICC\help\OICC_UM_sign-off-setup.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_finalisation.pdf", r"C:\Program Files (x86)\OICC\help\OICC_UM_other_stuff.pdf"]
        options2 = ['nothing', os.path.abspath(self.rootpath + r"\OICC_UM_PE.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_PE.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_PD.pdf"), os.path.abspath(
            self.rootpath + r"\OICC_UM_rectify_rejection.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_sign-off-setup.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_finalisation.pdf"), os.path.abspath(self.rootpath + r"\OICC_UM_other_stuff.pdf")]
        if evt_id < 8:
            if os.path.isfile(options[evt_id]):
                SP_Popen([options[evt_id]], shell=True)
            elif os.path.isfile(options[evt_id]):
                SP_Popen([options2[evt_id]], shell=True)
        else:
            if evt_id == 8:
                self.poplist()
                outs = os.path.abspath(self.rootpath + r'outstanding.html')
                if os.path.isfile(outs):
                    SP_Popen([outs], shell=True)
            elif evt_id == 9:
                rox = os.path.abspath(self.rootpath + r'\ROC')
                if os.path.exists(rox):
                    path = os.path.realpath(rox)
                    msg = wx.BusyInfo(
                        'Please wait, this will take some time to generate the files')
                    self.makeXML()
                    outs = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                pass
            elif evt_id == 10:
                if wordy.approvedpath == r'\\NT4\Client_Files\Public\PEOI\approved':
                    msg = wx.BusyInfo('Please wait, this will take some time')
                    self.xferttech(
                        r'\\AD2\Client_Files\Technical\PEOI\approved', wordy.approvedpath)
                else:
                    pass

            elif evt_id == 11:
                self.ROCxTOCxHTML()
                outs = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
                if os.path.isfile(outs):
                    SP_Popen([outs], shell=True)
                else:
                    pass

            elif evt_id == 12:
                outs = os.path.abspath(self.ROCpath + '\\' + "OPINDEX.csv")
                try:
                    with open(outs, 'wb') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                        wr.writerows(self.opindexgen())
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                    else:
                        pass
                except IOError:
                    print("cannot create opindex, someone has it open")
                    return

            elif evt_id == 13:
                outs = os.path.abspath(
                    self.ROCpath + '\\' + "detailed_outstanding.csv")
                try:
                    with open(outs, 'wb') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                        try:
                            x = self.det_out()
                            for y in x:
                                wr.writerows(y)
                        except Exception as e:
                            print('oops ' + str((inspect.stack()[0][2])))
                            print (e.message, e.args)
                            pass
                    if os.path.isfile(outs):
                        SP_Popen([outs], shell=True)
                    else:
                        pass
                except IOError:
                    print(
                        "cannot create detailed outstanding list, someone has it open")
                    return
            elif evt_id == 14:
                self.loggindb()
            elif evt_id == 15:
                self.reporter(event)
        return

    def XtracTor(self, target, pos):
        # it's the Vlookup again, bit fixed
        namez2 = []
        for nam3s in self.dbdata:
            temp = nam3s[0] + nam3s[1]
            if temp == target:
                temp2 = nam3s[pos]
                namez2.append(temp2)
        return namez2

    def RO_ROC(self, event):
        x = self.m_comboBox2qq.GetValue()
        if x != '':
            if x in self.m_comboBoxDB:
                if wordy.ROROC_frame_number == 0:
                    title = safestr(x)
                    frame = ROC_frame(title=title)
                    wordy.ROROC_frame_number = 1

    def God_Mode(self, event):
        x = self.m_comboBox2qq.GetValue()
        x2 = self.m_rev_master.GetValue()
        x3 = self.m_comboBox2TL.GetValue()
        if x != '' and x2 != '':
            if x in self.m_comboBoxDB:
                if x2 in self.tlist3:
                    if wordy.frame_number == 0:
                        wordy.extractedID = [safestr(x), safestr(x2),
                                             None]
                        wordy.do_save = 0
                        title = 'Record of Change God Mode'
                        frame = OtherFrame(title=title)
                        wordy.frame_number = 1
                        return
        elif x == '' and x2 == '' and x3 == '':
            if wordy.frame_number == 0:
                wordy.extractedID = ['', '',
                                     None]
                wordy.do_save = 0
                title = 'Record of Change God Mode'
                frame = OtherFrame(title=title)
                wordy.frame_number = 1

    def DBOK(self):
        if not os.path.isfile(wordy.DBpath):
            dlg = wx.MessageDialog(None, 'Cannot find database file!',
                                   '', wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            self.Close()
            return False
        else:
            return True

    def getnumfils(self):
        return len([name for name in os.listdir(self.DBpathARCH) if os.path.isfile(os.path.join(self.DBpathARCH, name))])

    def delOldest(self):
        oldest = min(os.listdir(self.DBpathARCH), key=lambda f: os.path.getctime(
            "{}/{}".format(self.DBpathARCH, f)))
        if fileaccesscheck(os.path.abspath(self.DBpathARCH + '\\' + oldest)):
            os.remove(os.path.abspath(self.DBpathARCH + '\\' + oldest))
        return

    def backUPtheDB(self):
        newest = max(os.listdir(self.DBpathARCH), key=lambda f: os.path.getctime(
            "{}/{}".format(self.DBpathARCH, f)))
        if (filecmp.cmp(wordy.DBpath, self.DBpathARCH + '\\' + newest, shallow=1)):
            return
        dumblist = [n for n in range(1, 99)]
        if self.getnumfils() > 0:
            for (dirpath, dirnames, fi) in os.walk(self.DBpathARCH):
                archfiles = []
                for fp in fi:
                    archfiles.append(
                        int(fp[-(len(fp) - (fp.rfind('.'))) + 1:]))
            newfile = 'ROC_db.pkl.' + \
                safestr(returnNotMatches(dumblist, archfiles)[0])
        else:
            newfile = 'ROC_db.pkl.1'
        try:
            shutil.copy(wordy.DBpath, os.path.abspath(
                self.DBpathARCH + '\\' + newfile))
            if self.getnumfils() > 10:
                while self.getnumfils() > 25:
                    if self.getnumfils() > 10:
                        self.delOldest()
                    else:
                        break
                rd = Path(self.DBpathARCH)
                while sum(f.stat().st_size for f in rd.glob('**/*') if f.is_file()) > 15000000:
                    if self.getnumfils() > 10:
                        self.delOldest()
                    else:
                        break
            else:
                return
        except Exception as e:
            print('oops ' + str((inspect.stack()[0][2])))
            print (e.message, e.args)
            pass
        return

    def det_out(self):
        valz = self.poplist(True)
        awQA = []
        awPE = []
        awPD = []
        rejd = []
        approved = []
        for x in valz:
            if isinstance(x, dict):
                PEstr = x.get('PE')
                PDstr = x.get('PD')
                QAstr = x.get('QA')
                RRstr = x.get('reject_reason')
                revstr = x.get('rev')
                PEOIstr = x.get('PEOI')
                DoCstr = x.get('doc')
                RfCstr = x.get('rfc')
                if isinstance(RRstr, list):
                    tt = safestr(''.join(RRstr))
                    RRstr = tt
                l1 = [PEstr, PDstr, QAstr]
                l2 = [awPE, awPD, awQA]
                for x, y in zip(l1, l2):
                    if x == 'rejected':
                        rejd.append(
                            [PEOIstr, revstr, DoCstr, RfCstr, RRstr])
                    elif x == 'awaiting':
                        if PEstr != 'rejected' and PDstr != 'rejected' and QAstr != 'rejected':
                            y.append(
                                [PEOIstr, revstr, DoCstr, RfCstr, RRstr])
                if PEstr == 'approved' and PDstr == 'approved' and QAstr == 'approved':
                    approved.append([PEOIstr, revstr, DoCstr, RfCstr, RRstr])
        inject = ["PEOI", "REV", "details of change",
                  "reason for change", "Comment"]
        inject2 = ["*****************", "*****",
                   "*****************", "*****************", "*****************"]

        valz4 = [
            [["No. in system: ", len(valz)]],
            [["Outstanding PE: ", len(awPE)]],
            [["Outstanding QA: ", len(awQA)]],
            [["Outstanding PD: ", len(awPD)]],
            [["Rejected: ", len(rejd)]],
            [["All Approved: ", len(approved)]],
            [inject2, ["Instructions Rejected"], [""], inject],
            rejd,
            [inject2, ["Instructions awaiting Prod Eng approval"], [""],  inject],
            awPE,
            [inject2, ["Instructions awaiting Production Manager approval"], [""],  inject],
            awPD,
            [inject2, ["Instructions awaiting Quality approval"], [""],  inject],
            awQA,
            [inject2, ["Instructions require issuing / finalising"], [""],  inject],
            approved
        ]

        return (valz4)

    def opindexgen(self):
        if self.DBOK():
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
        else:
            return
        PEOI = []
        for x in self.roctop:
            if isinstance(x, list):
                for y in x:
                    if isinstance(y, dict):
                        t = len(x)
                        if t > 1:
                            if isinstance(x[0], dict) and isinstance(x[-1], dict):
                                p1 = (x[0].get('details', {}).get('opno'))
                                p3 = (x[0].get('details', {}).get('obsolete', False))
                                p2 = (x[-1].get('revs', {}).get('rev'))
                                if p1 is not None and p2 is not None:
                                    PEOI.append(p1)
                                break
        PEOIx = sorted(PEOI)
        biglist = []
        biglist.append(["Project", "PEOI #", "Revision", "Description", "Implemented by",
                        "Date of revision*", "QA signed*", "Production signed*", "Date Received", "Details of Change", "Reason for Change", "Received by", "pages", "copies", "stage #s", "Pektron part numbers", "customer info", "link to file"])
        biglist.append(["*may not be availble (new feature May 2021)*"])
        for x in PEOIx:
            actpos = PEOI.index(x)
            opno = self.roctop[actpos][0].get('details', {}).get('opno')
            if x == opno:
                posy = len(self.roctop[actpos])
                if posy >= 1:
                    bl = self.verify_complete(actpos, posy - 1)
                    if bl:
                        biglist.append(bl)
        return(biglist)

    def verify_complete(self, actpos, cnt):
        opno = self.roctop[actpos][0].get('details', {}).get('opno')
        rev = self.roctop[actpos][cnt].get('revs', {}).get('rev')
        obstatus = self.roctop[actpos][0].get('details', {}).get('obsolete', False)
        if obstatus:
            obsval=" (PEOI marked obsolete)"
        else:
            obsval=""
        rev=rev+obsval
        opno=opno+obsval
        if int(self.roctop[actpos][cnt].get('revs', {}).get('PE_sign')) != 1:
            sig1 = 0
        else:
            sig1 = 1
        if int(self.roctop[actpos][cnt].get('revs', {}).get('PD_sign')) != 1:
            sig2 = 0
        else:
            sig2 = 1
        if int(self.roctop[actpos][cnt].get('revs', {}).get('QA_sign')) != 1:
            sig3 = 0
        else:
            sig3 = 1
        if sig1 + sig2 + sig3 == 3:
            impl = self.roctop[actpos][cnt].get('revs', {}).get('impl')
            desc = self.roctop[actpos][0].get('details', {}).get('stagedesc')
            OPINDEX_ISS = self.roctop[actpos][cnt].get(
                'revs', {}).get('rev_iss_date')
            OPINDEX_IMPL2 = self.roctop[actpos][cnt].get(
                'revs', {}).get('PE_sign_name')
            OPINDEX_qa = self.roctop[actpos][cnt].get(
                'revs', {}).get('QA_sign_name')
            OPINDEX_pd = self.roctop[actpos][cnt].get(
                'revs', {}).get('PD_sign_name')
            OPINDEX_date = self.roctop[actpos][cnt].get('revs', {}).get('date')
            OPINDEX_date = OPINDEX_date.replace(".", "/")
            if OPINDEX_IMPL2 is not None:
                impl = OPINDEX_IMPL2
            if OPINDEX_ISS is None:
                OPINDEX_ISS = "see PDF"
            if OPINDEX_qa is None:
                OPINDEX_qa = "see PDF"
            if OPINDEX_pd is None:
                OPINDEX_pd = "see PDF"
            OPINDEX_stageno = self.roctop[actpos][0].get(
                'details', {}).get('stageno')
            OPINDEX_pekpn = self.roctop[actpos][0].get(
                'details', {}).get('pekpn')
            OPINDEX_cstpn = self.roctop[actpos][0].get(
                'details', {}).get('cstpn')
            OPINDEX_copies = self.roctop[actpos][cnt].get(
                'revs', {}).get('copies')
            OPINDEX_rfc = self.roctop[actpos][cnt].get('revs', {}).get('rfc')
            OPINDEX_doc = self.roctop[actpos][cnt].get('revs', {}).get('doc')
            OPINDEX_rcvd = self.roctop[actpos][cnt].get('revs', {}).get('rcvd')
            OPINDEX_pages = self.roctop[actpos][cnt].get(
                'revs', {}).get('pages')
            out = safestr(opno + "_" + rev + ".pdf")
            proj = opno2proj(opno)
            appdirfil = ''
            link2 = "not found"
            if proj != '':
                appdirfil = os.path.abspath(wordy.approvedpath + '\\'
                                            + proj + '\\' + out)
            if os.path.isfile(appdirfil):
                link2 = appdirfil
            return[proj, opno, rev, desc, impl, OPINDEX_ISS, OPINDEX_qa, OPINDEX_pd, OPINDEX_date, OPINDEX_doc, OPINDEX_rfc, OPINDEX_rcvd, OPINDEX_pages, OPINDEX_copies, OPINDEX_stageno, OPINDEX_pekpn, OPINDEX_cstpn, link2]
        else:
            if cnt - 1 >= 1:
                return self.verify_complete(actpos, cnt - 1)
            else:
                return False
            

    def makeXML(self):
        if self.DBOK():
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
            rvs = []
            for  rec in progressbar(self.roctop):
                rvs = []
                savfil = safestr(rec[0]['details']['opno'])
                obstat = safestr(rec[0]['details'].get('obsolete', False))
                rvs.append({'details': {'opno': savfil}})
                rvs.append({'details': {'obsolete': obstat}})
                rvs.append(
                    {'details': {'cstpn': safestr(rec[0]['details']['cstpn'])}})
                rvs.append(
                    {'details': {'pekpn': safestr(rec[0]['details']['pekpn'])}})
                rvs.append(
                    {'details': {'desc': safestr(rec[0]['details']['desc'])}})
                rvs.append(
                    {'details': {'stageno': safestr(rec[0]['details']['stageno'])}})
                rvs.append(
                    {'details': {'stagedesc': safestr(rec[0]['details']['stagedesc'])}})
                found2 = foolproof_finder(rec, 1)
                if found2 > 0:
                    Fu = 1
                    while Fu <= found2:
                        rvs.append({'revs': {'rev': safestr(rec[Fu]['revs']['rev']), 'doc': safestr(rec[Fu]['revs']['doc']), 'rfc': safestr(rec[Fu]['revs']['rfc']), 'pages': safestr(rec[Fu]['revs']['pages']), 'copies': safestr(rec[Fu]['revs']['copies']), 'QA_sign': safestr(
                            rec[Fu]['revs']['QA_sign']), 'PE_sign': safestr(rec[Fu]['revs']['PE_sign']), 'PD_sign': safestr(rec[Fu]['revs']['PD_sign']), 'rcvd': safestr(rec[Fu]['revs']['rcvd']), 'impl': safestr(rec[Fu]['revs']['impl']), 'date': safestr(rec[Fu]['revs']['date'])}})
                        Fu += 1
                xml = dicttoxml.dicttoxml(rvs, attr_type=False)
                xml = xml.replace('<?xml version="1.0" encoding="UTF-8" ?>',
                                  '<?xml version="1.0" encoding="UTF-8" ?><?xml-stylesheet type="text/xsl" href="test2.xsl"?>')

                ROCsave = os.path.abspath(
                    self.ROCpath + '\\' + savfil + ".xml")
                ROCsave1 = os.path.abspath(
                    self.ROCpath + '\\' + "test2.xsl")
                ROCsave2 = os.path.abspath(
                    self.ROCpath + '\\' + savfil + ".html")               
                xmlfile = open(ROCsave, "w")
                xmlfile.write(xml.decode())
                xmlfile.close()
                xslt_doc = etree.parse(ROCsave1)
                xslt_transformer = etree.XSLT(xslt_doc)
                source_doc = etree.parse(ROCsave)
                output_doc = xslt_transformer(source_doc)
                output_doc.write(ROCsave2, pretty_print=True)
                self.ROCxTOCxHTML()

    def ROCxTOCxHTML(self):
        PagesHere = os.listdir(os.path.abspath(self.ROCpath))
        ROCsave3 = os.path.abspath(self.ROCpath + '\\' + "TOC.html")
        with open(ROCsave3, "w") as f:
            f.write(
                "<html><head><link rel='stylesheet' href='listscss.css'></head><body>\n")
            f.write("<h2>Printable Record of Change</h2><ul>")
            last = ''
            prange = 1
            jswitch = 0
            for filename in PagesHere:
                if filename.endswith(".html"):
                    if "TOC.html" not in filename:
                        sfnum, shortfile2, shortfile = self.dig_deep(filename)
                        if last == '':
                            while not (prange <= sfnum <= (prange + 99)):
                                prange = self.add_100(prange)
                                jswitch = 1
                        if not prange <= sfnum <= (prange + 99):
                            while prange < sfnum:
                                prange = self.add_100(prange)
                                if prange > sfnum:
                                    prange = prange - 100
                                    jswitch = 1
                                    break
                        if prange <= sfnum <= (prange + 99):
                            if last != '':
                                f.write("</ul></details>")
                                f.write("</ul></details>")
                            catt = str(str(prange) + " to " + str(prange + 99))
                            f.write(
                                "<details><summary>%s</summary><ul>" % (catt))
                            prange = self.add_100(prange)
                            jswitch = 1
                        if shortfile2 != last:
                            if last != '':
                                if jswitch == 0:
                                    f.write("</ul></details>")
                                elif jswitch == 1:
                                    jswitch = 0
                            f.write("<details><summary>%s</summary><ul>" %
                                    (shortfile2))
                        f.write("<li><a href='%s'>%s</a></li>" %
                                (filename, shortfile))
                        last = shortfile2
            f.write("</ul></details>")
            f.write("</ul></body></html>\n")

    def add_100(self, prange):
        prange = prange + 100
        return prange

    def dig_deep(self, filename):
        res = filename.rfind("PEOI")
        shortfile = filename[res:]
        shortfile2 = shortfile[5: 9]
        sfnum = int(shortfile2)
        return sfnum, shortfile2, shortfile

    def DBpoplist(self):
        if self.DBOK():
            tlist = []
            tlist2 = ['PEOI']
            tlist3=[(0,0)]
            tlist4 =[(0,0)]
            with open(wordy.DBpath, 'rb') as f:
                self.roctop = PKL_load(f)
            n = 1
            for x in self.roctop:
                tlist.append(x[0]['details']['opno'])
                (fx, fx2) = self.find_proj(x[0]['details']['opno'])
                if fx not in tlist2:
                    tlist2.append(fx)
                    tlist3.append((fx, n))
                    try:
                        tlist4.append((int(fx), n))
                    except:
                        tlist4.append((999999999, n))
                    n+=1
            tlist4.sort()        
            tlist5=sorted(tlist4, reverse=True)
            tlist5.pop()
            tlist5.insert(0, (0,0))
            if self.sortorder.GetValue():
                masterlist = tlist4
            else:
                masterlist = tlist5
            list5=[None for _ in range(len(masterlist))]
            for xx in range(len(masterlist)):
                list5[xx]=tlist3[masterlist[xx][1]][0]
            list5[0]="PEOI"
            return (tlist, list5)

