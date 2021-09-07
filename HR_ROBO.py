import copy
import sys
import pandas
from pubsub import pub
import pyautogui
import wx
import wx.lib.scrolledpanel
from sheetdata import *


class PromptingComboBox(wx.ComboBox) :
    def __init__(self, parent, choices=[], style=0, **par):
        wx.ComboBox.__init__(self, parent, wx.ID_ANY, style=style|wx.CB_DROPDOWN, choices=choices, **par)
        self.choices = choices
        self.Bind(wx.EVT_TEXT, self.OnText)
        self.Bind(wx.EVT_KEY_DOWN, self.OnPress)
        self.ignoreEvtText = False
        self.deleteKey = False

    def OnPress(self, event):
        if event.GetKeyCode() == 8:
            self.deleteKey = True
        event.Skip()

    def OnText(self, event):
        currentText = event.GetString()
        if self.ignoreEvtText:
            self.ignoreEvtText = False
            return
        if self.deleteKey:
            self.deleteKey = False
            if self.preFound:
                currentText =  currentText[:-1]

        self.preFound = False
        for choice in self.choices :
            if choice.startswith(currentText):
                self.ignoreEvtText = True
                self.SetValue(choice)
                self.SetInsertionPoint(len(currentText))
                self.SetTextSelection(len(currentText), len(choice))
                self.preFound = True
                break

class Checker_Process(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, "HR ROBO", size=(420, 310))
        self.panel = wx.Panel(self)
        try:
            image_file = os.path.join(UTIL_PATH,'LOGO.jpg')
            # image_file = os.path.join(UTIL_PATH,'sslogo2.jpeg')
            bmp1 = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.bitmap1 = wx.StaticBitmap(self.panel, -1, bmp1, (0, 0))
        except:
            pass

        self.label = wx.StaticText(self.panel, -1, "", (10, 250))
        self.labeltext = wx.StaticText(self.panel, -1, "Powered by SequelString", (265, 250))
        self.start_button = wx.Button(self.panel, 1, label="Show All Employees", pos=(110, 160), size=(180, 60))
        self.Bind(wx.EVT_BUTTON, self.start_f, self.start_button)
        font = wx.SystemSettings.GetFont(wx.SYS_SYSTEM_FONT)

        font.SetPointSize(20)
        self.start_button.SetBackgroundColour("#0B0B3B")
        self.start_button.SetFont(font)
        self.start_button.SetForegroundColour("#FAFAFA")


    def start_f(self,event):
        HRROBO(None, title='HR-ROBO')


class LoginDialog(wx.Dialog):
    """
    Class to define login dialog
    """

    # ----------------------------------------------------------------------
    def __init__(self):
        """Constructor"""
        wx.Dialog.__init__(self, None, title="Login")
        # user info
        user_sizer = wx.BoxSizer(wx.HORIZONTAL)
        user_lbl = wx.StaticText(self, label="Username:")
        user_sizer.Add(user_lbl, 0, wx.ALL | wx.CENTER, 5)
        self.user = wx.TextCtrl(self)
        user_sizer.Add(self.user, 0, wx.ALL, 5)

        # pass info
        p_sizer = wx.BoxSizer(wx.HORIZONTAL)
        p_lbl = wx.StaticText(self, label="Password:")
        p_sizer.Add(p_lbl, 0, wx.ALL | wx.CENTER, 5)
        self.password = wx.TextCtrl(self, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER)
        p_sizer.Add(self.password, 0, wx.ALL, 5)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(user_sizer, 0, wx.ALL, 5)
        main_sizer.Add(p_sizer, 0, wx.ALL, 5)

        btn = wx.Button(self, label="Login")
        btn.Bind(wx.EVT_BUTTON, self.onLogin)
        main_sizer.Add(btn, 0, wx.ALL | wx.CENTER, 5)
        self.SetSizer(main_sizer)

    # ----------------------------------------------------------------------
    def onLogin(self, event):
        """
        Check credentials and login
        """
        creds = {"Hello1": "1234", "admin": "admin", "user3": "12345"}
        user_name = self.user.GetValue()
        user_password = self.password.GetValue()
        try:
            getuser = creds[user_name]
        except KeyError:
            getuser = None
        if getuser:
            if creds[getuser] != user_password:
                pass
            else:
                pub.sendMessage("frameListener", message="show")
                self.Destroy()


class HRROBO(wx.Frame):

    def __init__(self, parent, title):
        # pub.subscribe(self.myListener, "frameListener")

        self.get_users()
        # self.data = {'Timestamp': '6/13/2019 22:52:50', 'Salutation/Title': 'Mr.', 'First Name': 'Rahul', 'Middle Name': '', 'Last Name': 'Jain', 'Gender': 'Male', 'Father’s/Husband Name': 'Arun Jain', 'Employee Relation': 'Father', 'Employee Name': 'Rahul Jain', 'Marital Status': 'Married', 'Spouse Name': 'Anjali', 'No. of Children': 0, 'Present Address': 'h-123, kuch bhi', 'Present City': 'pune', 'Present State': 'Maharashtra', 'Present Pin code': 343434, 'Present Phone': 8979878979, 'Permanent-Address': '', 'Permanent Address': 'h-456, or kuch', 'Permanent City': 'ghaziabad', 'Permanent State': 'Uttar Pradesh', 'Permanent Pin code': 201002, 'Permanent Phone': 9878799799, 'Primary Bank Name': 'ohyes', 'Primary IFSC Code': 'ohho12345', 'Primary Account No': 54344343, 'Primary name as per bank': 'Rahul', 'MICR': 'lsdnlsn', 'Expected Date of Joining': '9/9/2019', 'Date of Birth': '9/9/1998', 'Qualification': 'Graduate', 'Blood Group': 'b+', 'Emergency Phone No': 98797979999, 'Emergency Contact Person': 'Ravi jain', 'Aadhar card no': 986987979896, 'Permanent Account No': 678979879, 'Email Address': 'rahul.rahul@yahoo.com', 'Personal Mobile No': 98797897987, 'Choose One': 'Yes', 'Establishment Address': 'h678, bolo bolo', 'Universal Account Number': 8979879, 'PF Account Number': 79879879, 'Date of joining (Unexempted)': '9/9/2019', 'Date of exit (Unexempted)': '9/9/2019', 'Scheme Certificate (if issued)': 798798, 'PPO Number (if issued)': 768788, 'Non-Contributory Period (NCP) Days': 89, 'Name & Address of the Trust': '', 'UAN': '', 'Member EPS A/c No': '', 'Date of joining (Exempted)': '', 'Date of exit (Exempted)': '', 'Scheme Certificate No (if issued)': '', 'Non Contributory period (NCP) Days': '', 'Worked': 'Yes', 'State Country of origin': 'videsh', 'Passport No': 9879879879, 'Validity of Passport': '5/30/2029', 'Upload Passport Photo': 'https://drive.google.com/open?id=1-0X7L6f-GqOVboY1vIJUfLHRZr9Cp4gE', 'Upload Docs': '', 'Employee Number': 'hello', 'Confirm Date of Joining': '', 'Training From': '', 'Date of Probation': '', 'Date of Retirement': '', 'Date of confirmation': '', 'Payroll Code': '', 'Category Code': '', 'Status Code': '', 'Grade Code': '', 'Designation': '', 'Cost Center Code': '', 'Business Area Code': 'Kandivali', 'Location Code': '', 'Sub-Location Code': '', 'Occupation Code': '', 'PF Registration Code': 'KDMAL0213914', 'PF wef Dt': '', 'E.S.I. No.': '', 'Reports to emp': '', "Reporting Manager's Token ID": '', 'Mobile number': '', 'Web user name': 'hello', 'Web user password': '', 'Web access level': 0, 'Attendance user id': '', 'Designation Code': '', 'Department Code': '', 'Sub-Department Code': 'SCM-Sales', 'CR Mem No': '', 'Client name': '', 'Client name code': '', 'State code': '', 'Group joining date': ''}



        # self.box.Enable()

        # Ask user to login
        # dlg = LoginDialog()
        # dlg.ShowModal()
        super(HRROBO, self).__init__(parent, title=title, size=(1100, 700))
        self.InitUI()
        self.Centre()
        self.Show()

    def myListener(self, message, arg2=None):
        """
        Show the frame
        """
        self.Show()

    def get_users(self):
        dict_data, list_data,self.codedict = get_records()
        all_names = {}
        tmpval = []
        for i, x in enumerate(list_data[1:]):
            name = "{} {} {}".format(str(x[2]), str(x[3]), str(x[4]))
            all_names[name] = i
        #     tmpval.append(f"{x[34].strip()}{x[35].strip()}")
        # if len(tmpval) != len(set(tmpval)):
        #     res = pyautogui.alert("Multiple entries found with the same details. \nPlease Check in Google Sheet\n\n\nPress OK to Abort.")
        #     sys.exit()

        self.box = wx.SingleChoiceDialog(None, "Please Select a User You want to continue with", "Select Type",
                                    list(all_names.keys()))
        if self.box.ShowModal() == wx.ID_OK:
            ans = self.box.GetStringSelection()
        else:
            al = pyautogui.alert(
                "Please Select a user to continue.\n\n\nPress OK to Abort.")
            sys.exit()
        self.box.Disable()
        self.ind = int(all_names.get(ans))
        self.data = dict_data[all_names.get(ans)]
        #####################
        #
        # self.data ={'Timestamp': '7/3/2019 0:27:26', 'Salutation/Title': 'Mr.', 'First Name': 'Sahil', 'Middle Name': '', 'Last Name': 'Kasana', 'Gender': 'Male', 'Father’s/Husband Name': 'Manveer', 'Employee Relation': 'Father', 'Employee Name': 'Sahil Kasana', 'Marital Status': 'Unmarried', 'Spouse Name': '', 'No. of Children': '', 'Present Address': 'H.no Pna', 'Present City': 'Noida', 'Present State': 'UP', 'Present Pin code': 201301, 'Present Phone': 9777766667767, 'Permanent-Address': 'Same As Above', 'Permanent Address': 'H.no Pna', 'Permanent City': 'Noida', 'Permanent State': 'UP', 'Permanent Pin code': 201301, 'Permanent Phone': 9777766667767, 'Primary Bank Name': 'yes bank', 'Primary IFSC Code': 'KKDF0005019', 'Primary Account No': 324332, 'Primary name as per bank': 'Sahil Kasan', 'MICR': 123456789, 'Expected Date of Joining': '10/19/2019', 'Date of Birth': '9/9/2019', 'Qualification': 'b tech', 'Blood Group': 'B-', 'Emergency Phone No': 87887878787, 'Emergency Contact Person': 'Khatana', 'Aadhar card no': 88977899887979, 'Permanent Account No': 9797987979, 'Email Address': 'cutesahil@gmial.com', 'Personal Mobile No': 988989, 'Choose One': 'Un-exempted', 'Establishment Address': 'lll', 'Universal Account Number': 8880, 'PF Account Number': 98898, 'Date of joining (Unexempted)': '8/8/2019', 'Date of exit (Unexempted)': '8/8/2019', 'Scheme Certificate (if issued)': 'uhuuh', 'PPO Number (if issued)': 'uhhuh', 'Non-Contributory Period (NCP) Days': 88, 'Name & Address of the Trust': '', 'UAN': '', 'Member EPS A/c No': '', 'Date of joining (Exempted)': '', 'Date of exit (Exempted)': '', 'Scheme Certificate No (if issued)': '', 'Non Contributory period (NCP) Days': '', 'Worked': 'No', 'State Country of origin': '', 'Passport No': '', 'Validity of Passport': '', 'Upload Passport Photo': 'https://drive.google.com/open?id=1ObKyIzInmnYyIJlBDJQKcspiv98Ebq-F', 'Upload Docs': 'https://drive.google.com/open?id=10qMC8l769zcmU8H6iNAAlaoa42pp6364', '': 'cutesahil@gmial.com', 'Employee Number': 'S321009', 'Confirm Date of Joining': '', 'Training From': '', 'Date of Probation': '', 'Date of Retirement': '', 'Date of confirmation': '', 'Payroll Code': '', 'Category Code': '', 'Status Code': '', 'Grade Code': '', 'Designation': '', 'Cost Center Code': '', 'Business Area Code': '', 'Location Code': '', 'Sub-Location Code': '', 'Occupation Code': '', 'PF Registration Code': 'KDMAL0213914000', 'PF wef Dt': '', 'E.S.I. No.': '', 'Reports to emp': '', "Reporting Manager's Token ID": '', 'Mobile number': '', 'Web user name': 'S321009', 'Web user password': '', 'Web access level': 0, 'Attendance user id': '', 'Designation Code': '', 'Department Code': 'Farm Division- Staffing(FDSTAFF)', 'Sub-Department Code': 'SCM-Backend(22)', 'CR Mem No': '', 'Client name': '', 'Client name code': 'Mahindra Home Finance (19)', 'State code': 'MIZORAM(MZ)', 'Group joining date': '', 'Reporting Manager Email id': ''}
        # self.codedict = OrderedDict([('S', 321008), ('S1S', 21000), ('S2S', 21000), ('S3S', 21000), ('S4S', 21000), ('S5S', 21000), ('S6S', 21000), ('S7S', 21000), ('S8S', 21000), ('S9S', 21000), ('S1M', 21000), ('S2M', 21000), ('S3M', 21000), ('9', 321000), ('NA', 21001)])


    def InitUI(self):
        self.p = wx.lib.scrolledpanel.ScrolledPanel(self, -1, size=(830, 600), pos=(0, 0),
                                                    style=wx.SIMPLE_BORDER)

        self.p.SetupScrolling()
        # self.p.SetBackgroundColour('#DAF7A6')
        self.p.SetBackgroundColour('GREY')

        # data = {'Timestamp': '6/13/2019 1:39:57', 'First Name': 'john', 'Middle Name': '', 'Last Name': 'dscd', 'Gender': 'Male', 'Father’s/Husband Name': ' kkmkmk', 'Employee Relation': 'kjnkjn', 'Employee Name': 'lklm', 'Marital Status': 'Unmarried', 'Spouse Name': 'NA', 'No. of Children': 0, 'Present Address': 'dsklmdslk dkkmckm ', 'Present City': 'ksmskc', 'Present State': 'c dk ', 'Present Pin code': 98098098, 'Present Phone': 9890899898, 'Permanent Address': 'kbjnjnklmnl kknlknlk', 'Permanent City': 'lknlkn', 'Permanent State': 'lknln', 'Permanent Pin code': 87987979, 'Permanent Phone': 8799779797, 'Primary Bank Name': 'iuhk', 'Primary IFSC Code': 'kjhkjh', 'Primary Account No': 8986987897, 'Primary name as per bank': 'chg', 'MICR': 'kkjnkjnl', 'Expected Date of Joining': '8/9/2019', 'Date of Birth': '7/8/1998', 'Qualification': 'JNNN', 'Blood Group': 'KNLKL', 'Emergency Phone No': 898799797, 'Emergency Contact Person': 'LJLJL', 'Emergency Details': '', 'Aadhar card No': 87897976797, 'Email id': 8779879, 'Personal Mobile No': 987897987, 'Un-exempted': 'Yes', 'Establishment Address': 'wlclkdsnl', 'Universal Account Number': 8897979, 'PF Account Number': 78, 'Date of joining (MM/DD/ YYYY)': '6/6/2019', 'Date of exit (MM/DD/ YYYY)': '6/6/2019', 'Scheme Certificate (if issued)': 'Yes', 'PPO Number (if issued)': 8798, 'Non Contributory Period (NCP)': 87, 'Exempted Trusts': 'Yes', 'Name & Address of the Trust': 'Nkhkjh', 'UAN ': 'hj', 'Member EPS A/c No ': 'khjh', 'Date of joining (MM/DD/YYYY)': '6/5/2019', 'Date of exit (MM/DD/ YYYY) ': '4/6/2019', 'Scheme Certificate No (if issued)': 'Yes', 'Non Contributory period  (NCP)': 987, 'Worked': 'Yes', 'State Country of origin ': 'iuoiuiou', 'Passport No': 'oiuoi', 'Validity of Passport': '8/8/2019', 'Upload': 'https://drive.google.com/open?id=1BZ7YRR71k_st5iH2ohnIe3I1xv4zSytG', 'Permanent Account No': 8797, 'Salutation/Title': 'Mr.'}
        # data = {'Timestamp': '6/13/2019 1:39:57',"Salutation/Title":"101","First Name":"102","Middle Name":"103","Last Name":"104","Gender":"105","Father’s/Husband Name":"106","Employee Relation":"107","Employee Name":"108","Marital Status":"109","Spouse Name":"110","No. of Children":"111","Present Address":"112","Present City":"113","Present State":"114","Present Pin code":"115","Present Phone":"116","Permanent Address":"117","Permanent City ":"118","Permanent State":"119","Permanent Pin code":"120","Permanent Phone":"121","Primary Bank Name":"122","Primary IFSC Code":"123","Primary Account No":"124","Primary name as per bank":"125","MICR":"126","Expected Date of Joining":"127","Date of Birth":"128","Qualification":"129","Blood Group":"130","Emergency Phone No":"131","Emergency Contact Person":"132","Aadhar card no":"133","Permanent Account No":"134","Email Address":"135","Personal Mobile No":"136","Establishment Address ":"137","Universal Account Number ":"138","PF Account Number ":"139","Date of joining (Unexempted) ":"140","Date of exit (Unexempted) ":"141","Scheme Certificate (if issued) ":"142","PPO Number (if issued) ":"143","Non-Contributory Period (NCP) Days":"144","Name & Address of the Trust ":"145","UAN ":"146","Member EPS A/c No ":"147","Date of joining (Exempted) ":"148","Date of exit (Exempted) ":"149","Scheme Certificate No (if issued) ":"150","Non Contributory period (NCP) Days":"151","State Country of origin ":"152","Passport No. ":"153","Validity of Passport ":"154","Upload Passport Photo":"155","Upload Docs":"156","Employee Number":"157","Confirm Date of Joining":"158","Training From":"159","Date of Probation":"160","Date of Retirement":"161","Date of confirmation":"162","Payroll Code":"163","Category Code ":"164","Status Code":"165","Grade Code":"166","Designation":"167","Cost Center Code":"168","Business Area Code":"169","Location Code":"170","Occupation Code":"171","PF Registration Code":"172","PF wef Dt":"173","E.S.I. No.":"174","Reports to emp":"175","Reporting Manager's Token ID ":"176","Mobile number":"177","Web user name":"178","Web user password":"179","Web access level":"180","Leave policy ID":"181","Attendance user id":"182","Designation Code":"183","Department Code":"184","CR Mem No":"185","Client name":"186","Client name code":"187","State code":"188","Group joining date":"189"}

        mainSizer = wx.BoxSizer(wx.HORIZONTAL)

        gs = wx.FlexGridSizer(50, 4, 5, 15)
        self.label1 = wx.StaticText(self.p, label="Salutation/Title:")
        self.label2 = wx.StaticText(self.p, label="First Name:")
        self.label3 = wx.StaticText(self.p, label="Middle Name:")
        self.label4 = wx.StaticText(self.p, label="Last Name:")
        self.label5 = wx.StaticText(self.p, label="Gender:")
        self.label6 = wx.StaticText(self.p, label="Father’s/Husband Name:")
        self.label7 = wx.StaticText(self.p, label="Employee Relation:")
        self.label8 = wx.StaticText(self.p, label="Employee Name:")
        self.label9 = wx.StaticText(self.p, label="Marital Status:")
        self.label10 = wx.StaticText(self.p, label="Spouse Name:")
        self.label11 = wx.StaticText(self.p, label="No. of Children:")
        self.label12 = wx.StaticText(self.p, label="Present Address:")
        self.label13 = wx.StaticText(self.p, label="Present City:")
        self.label14 = wx.StaticText(self.p, label="Present State:")
        self.label15 = wx.StaticText(self.p, label="Present Pin code:")
        self.label16 = wx.StaticText(self.p, label="Present Phone:")
        self.label17 = wx.StaticText(self.p, label="Permanent Address:")
        self.label18 = wx.StaticText(self.p, label="Permanent City :")
        self.label19 = wx.StaticText(self.p, label="Permanent State:")
        self.label20 = wx.StaticText(self.p, label="Permanent Pin code:")
        self.label21 = wx.StaticText(self.p, label="Permanent Phone:")
        self.label22 = wx.StaticText(self.p, label="Primary Bank Name:")
        self.label23 = wx.StaticText(self.p, label="Primary IFSC Code:")
        self.label24 = wx.StaticText(self.p, label="Primary Account No:")
        self.label25 = wx.StaticText(self.p, label="Primary name as per bank:")
        self.label26 = wx.StaticText(self.p, label="MICR:")
        self.label27 = wx.StaticText(self.p, label="Expected Date of Joining(MM/DD/YYYY):")
        self.label28 = wx.StaticText(self.p, label="Date of Birth:(MM/DD/YYYY)")
        self.label29 = wx.StaticText(self.p, label="Qualification:")
        self.label30 = wx.StaticText(self.p, label="Blood Group:")
        self.label31 = wx.StaticText(self.p, label="Emergency Phone No:")
        self.label32 = wx.StaticText(self.p, label="Emergency Contact Person:")
        self.label33 = wx.StaticText(self.p, label="Aadhar card no:")
        self.label34 = wx.StaticText(self.p, label="Permanent Account No:")
        self.label35 = wx.StaticText(self.p, label="Email Address:")
        self.label36 = wx.StaticText(self.p, label="Personal Mobile No:")
        self.label37 = wx.StaticText(self.p, label="Establishment Address :")
        self.label38 = wx.StaticText(self.p, label="Universal Account Number :")
        self.label39 = wx.StaticText(self.p, label="PF Account Number :")
        self.label40 = wx.StaticText(self.p, label="Date of joining (Unexempted):")
        self.label41 = wx.StaticText(self.p, label="Date of exit (Unexempted) :")
        self.label42 = wx.StaticText(self.p, label="Scheme Certificate (if issued) :")
        self.label43 = wx.StaticText(self.p, label="PPO Number (if issued) :")
        self.label44 = wx.StaticText(self.p, label="Non-Contributory Period (NCP) Days:")
        self.label45 = wx.StaticText(self.p, label="Name & Address of the Trust :")
        self.label46 = wx.StaticText(self.p, label="UAN :")
        self.label47 = wx.StaticText(self.p, label="Member EPS A/c No :")
        self.label48 = wx.StaticText(self.p, label="Date of joining (Exempted) :")
        self.label49 = wx.StaticText(self.p, label="Date of exit (Exempted) :")
        self.label50 = wx.StaticText(self.p, label="Scheme Certificate No (if issued) :")
        self.label51 = wx.StaticText(self.p, label="Non Contributory period (NCP) Days:")
        self.label52 = wx.StaticText(self.p, label="State Country of origin :")
        self.label53 = wx.StaticText(self.p, label="Passport No. :")
        self.label54 = wx.StaticText(self.p, label="Validity of Passport :")
        self.label55 = wx.StaticText(self.p, label="Upload Passport Photo:")
        self.label56 = wx.StaticText(self.p, label="Upload Docs:")
        self.label57 = wx.StaticText(self.p, label="Employee Number:")
        self.label58 = wx.StaticText(self.p, label="Confirm Date of Joining(MM/DD/YYYY):")
        self.label59 = wx.StaticText(self.p, label="Training From:")
        self.label60 = wx.StaticText(self.p, label="Date of Probation(MM/DD/YYYY):")
        self.label61 = wx.StaticText(self.p, label="Date of Retirement(MM/DD/YYYY):")
        self.label62 = wx.StaticText(self.p, label="Date of confirmation(MM/DD/YYYY):")
        self.label63 = wx.StaticText(self.p, label="Payroll Code:")
        self.label64 = wx.StaticText(self.p, label="Category Code :")
        self.label65 = wx.StaticText(self.p, label="Status Code:")
        self.label66 = wx.StaticText(self.p, label="Grade Code:")
        self.label67 = wx.StaticText(self.p, label="Designation:")
        self.label68 = wx.StaticText(self.p, label="Cost Center Code:")
        self.label69 = wx.StaticText(self.p, label="Business Area Code:")
        self.label70 = wx.StaticText(self.p, label="Location Code:")
        self.label71 = wx.StaticText(self.p, label="Sub-Location Code:")
        self.label72 = wx.StaticText(self.p, label="Occupation Code:")
        self.label73 = wx.StaticText(self.p, label="PF Registration Code:")
        self.label74 = wx.StaticText(self.p, label="PF wef Dt:")
        self.label75 = wx.StaticText(self.p, label="E.S.I. No.:")
        self.label76 = wx.StaticText(self.p, label="Reports to emp:")
        self.label77 = wx.StaticText(self.p, label="Reporting Manager's Token ID :")
        self.label78 = wx.StaticText(self.p, label="Mobile number:")
        self.label79 = wx.StaticText(self.p, label="Web user name:")
        self.label80 = wx.StaticText(self.p, label="Web user password:")
        self.label81 = wx.StaticText(self.p, label="Web access level:")
        self.label82 = wx.StaticText(self.p, label="Attendance user id:")
        self.label83 = wx.StaticText(self.p, label="Designation Code:")
        self.label84 = wx.StaticText(self.p, label="Department Code:")
        self.label85 = wx.StaticText(self.p, label="Sub-Department Code:")
        self.label86 = wx.StaticText(self.p, label="CR Mem No:")
        self.label87 = wx.StaticText(self.p, label="Client name:")
        self.label88 = wx.StaticText(self.p, label="Client name code:")
        self.label89 = wx.StaticText(self.p, label="State code:")
        self.label90 = wx.StaticText(self.p, label="Group joining date(MM/DD/YYYY):")
        self.label91 = wx.StaticText(self.p, label="Reporting Manager Email id")
        ###########################################################################
        # self.text1 = wx.TextCtrl(self.p)
        self.codeseries = {"01": "S", "02": "S", "03": "S", "04": "S", "05": "S1S", "06": "S1M", "07": "S", "08": "S", "09": "S", "10": "S", "100": "S", "101": "S", "102": "S", "103": "S", "104": "S", "105": "S", "109": "S", "11": "S", "110": "S3S", "111": "S3S", "112": "S", "113": "S", "114": "S", "115": "S", "12": "S", "121": "S", "123": "S", "124": "S", "125": "S3S", "126": "S", "13": "S", "14": "S", "15": "S", "16": "S", "17": "S4S", "18": "S", "19": "S", "20": "9", "21": "S", "22": "S", "23": "S", "24": "S3S", "25": "S2S", "26": "S", "27": "S2S", "28": "S", "29": "S", "30": "S", "31": "S", "32": "S6S", "33": "S", "34": "S", "35": "S", "36": "S7S", "37": "S5S", "38": "S", "39": "S", "40": "S", "41": "S", "42": "S", "43": "S", "44": "S", "45": "S", "46": "S", "47": "S", "48": "S", "49": "S", "50": "S", "51": "S", "52": "S", "53": "S", "54": "S", "55": "S", "56": "S", "57": "S3S", "58": "S", "59": "S", "60": "S", "61": "S", "62": "S", "63": "S", "64": "S", "65": "S", "66": "S", "67": "S", "68": "S", "69": "S", "70": "S", "71": "S", "72": "S", "73": "S", "74": "S", "75": "S", "76": "S", "77": "S8S", "78": "S", "79": "S2M", "80": "S", "81": "S", "82": "S", "83": "S8S", "84": "S6S", "85": "S1S", "86": "S", "87": "S", "88": "S", "89": "S", "90": "S6S", "91": "S", "92": "S1S", "93": "S", "94": "S", "95": "S", "96": "S2M", "97": "S", "98": "S", "99": "S","129":"S"}
        self.chtext1 = OrderedDict({"Mr.": "Mr.", "Ms.": "Ms.", "Mrs.": "Mrs."})
        self.chtext63 = OrderedDict([('01', '1'), ('02', '2'), ('03', '3'), ('04', '4'), ('05', '5'), ('06', '6'), ('07', '7'), ('08', '8'), ('09', '9')])
        self.chtext64 = OrderedDict([('01', 'GENERAL'), ('02', 'CONTRACT')])
        self.chtext65 = OrderedDict([('01', 'Permanent'), ('02', 'Probation'), ('03', 'Trainee'), ('08', 'Contract')])
        self.chtext66 = OrderedDict([("CONT", "CONTRACT"),("EX", "Executive"),("GR", "Default"),("L1", "L1"),("L10O", "L10 - Operational"),("L3", "L3"),("L9O", "L9 - Operational"),("MB1", "MB1"),("MB2", "MB2"),("MS0", "MS0"),("MS1", "MS1"),("MS2", "MS2"),("MS3", "MS3"),("MS4", "MS4"),("MT1", "MT1"),("MT2", "MT2"),("MT3", "MT3")])
        self.chtext68 = OrderedDict(
            [('00', 'Default'), ('01', 'MIBL'), ('06', 'Common'), ('1', 'MIBS.STF.MM.SBUD.KND.00001012'),
             ('27', 'MBPO Accounts'), ('MB10381002', 'FES Salary'), ('MB10381003', 'AS Haridwar(Salary)'),
             ('MB10381004', 'Rudrapur Salary'), ('MB10381010', 'MB10381010'), ('MB10381013', 'Chakan Salary'),
             ('MB10381050', 'Cenvat-M&M FES'), ('MB10381051', 'KND BP M&M AS'), ('MB10381052', 'KND BP M&M FES'),
             ('MB10381056', 'FES Freight Bill passing'), ('MB10381059', 'MADPL Accounting Activity'),
             ('MB10381061', 'Nashik Salary'), ('MB10381063', 'Nashik- CNV M&M AS'),
             ('MB10381064', 'Nashik Bill passing'), ('MB10381067', 'MADPL Import'),
             ('MB10381068', 'Nashik- Accounting Activities'), ('MB10381069', 'Igatpuri Salary'),
             ('MB10381072', 'Igatpuri Bill passing'), ('MB10381073', 'Zaheerabad Salary'),
             ('MB10381078', 'Haridwar Bill passing'), ('MB10381080', 'Nagpur Salary'),
             ('MB10381083', 'Nagpur Bill passing'), ('MB10381085', 'Bill passing  M&M FES'),
             ('MB10381086', 'Accounting Activities M&M FES'), ('MB10381087', 'Jaipur BP M&M SBU'),
             ('MB10381088', 'Kanhe Cenvat'), ('MB10381089', 'Wadgaon Bill passing'),
             ('MB10381091', 'Chakan Cenvat Bill passing'), ('MB10381092', 'Chakan Bill Passing'),
             ('MB10381093', 'Chakan Freight Bill passing'), ('MB10381095', 'HR/Admin'),
             ('MB10381099', 'Chennai Bill passing'), ('MB10381100', 'MB10381100'), ('MB10381101', 'Salary Common'),
             ('MB10381102', 'Bill Passing Common'), ('MB10381103', 'General Overhead'), ('MB10381118', 'Chakan Common'),
             ('MB10381119', 'Nashik Common'), ('MB10381120', 'Igatpuri Common'), ('MB10381121', 'Nagpur Common'),
             ('MB10381122', 'Zaheerabad Common'), ('MB10381123', 'Worli Common'), ('MB10381124', 'Kandivali Common'),
             ('MB10381125', 'Staffing '), ('MB10381204', 'Solapur'), ('MB10381221', 'HRM - Manpower Service'),
             ('MB10382095', 'MB10382095'), ('MIBS.STF.GC.1049.HYD', 'MIBS.STF.GC.1049.HYD.052'),
             ('MIBS.STF.GC.1049.KSH', 'MIBS.STF.GC.1049.KSH.052'), ('MIBS.STF.GC.1049.PTN', 'MIBS.STF.GC.1049.PTN.052'),
             ('MIBS.STF.GC.1049.RPR', 'MIBS.STF.GC.1049.RPR.052'), ('MIBS.STF.GC.9025.PNE', 'MIBS.STF.GC.9025.PNE.001'),
             ('MIBS.STF.MM.AUTO.KOL', 'MIBS.STF.MM.AUTO.KOL.060'), ('MIBS.STF.MM.AUTO.NSK', 'MIBS.STF.MM.AUTO.NSK.001'),
             ('MIBS.STF.MM.FESD.KOL', 'MIBS.STF.MM.FESD.KOL.051'), ('MIBS.STF.MM.FESD.LKW', 'MIBS.STF.MM.FESD.LKW.051'),
             ('STF.EX.BHRC.CLB.001', 'STF.EX.BHRC.CLB.001'), ('STF.EX.DNB1.KND.050', 'STF.EX.DNB1.KND.050'),
             ('STF.EX.FK01.BLR.050', 'STF.EX.FK01.BLR.050'), ('STF.EX.HOUZ.BLR.010', 'STF.EX.HOUZ.BLR.010'),
             ('STF.GC.1007.ADH.003', 'STF.GC.1007.ADH.003'), ('STF.GC.1007.ADH.007', 'STF.GC.1007.ADH.007'),
             ('STF.GC.1007.ADH.010', 'STF.GC.1007.ADH.010'), ('STF.GC.1007.BSR.004', 'STF.GC.1007.BSR.004'),
             ('STF.GC.1007.CHN.005', 'STF.GC.1007.CHN.005'), ('STF.GC.1007.KND.002', 'STF.GC.1007.KND.002'),
             ('STF.GC.1007.PMR.001', 'STF.GC.1007.PMR.001'), ('STF.GC.1007.THN.006', 'STF.GC.1007.THN.006'),
             ('STF.GC.1007.THN.009', 'STF.GC.1007.THN.009'), ('STF.GC.1007.THN.010', 'STF.GC.1007.THN.010'),
             ('STF.GC.1007.WRL.008', 'STF.GC.1007.WRL.008'), ('STF.GC.1007.WRL.010', 'STF.GC.1007.WRL.010'),
             ('STF.GC.1007.WRL.011', 'STF.GC.1007.WRL.011'), ('STF.GC.1008.CHN.001', 'STF.GC.1008.CHN.001'),
             ('STF.GC.1009.JPR.001', 'STF.GC.1009.JPR.001'), ('STF.GC.1011.CHN.001', 'STF.GC.1011.CHN.001'),
             ('STF.GC.1012.KND.050', 'STF.GC.1012.KND.050'), ('STF.GC.1012.NGR.001', 'STF.GC.1012.NGR.001'),
             ('STF.GC.1012.NGR.050', 'STF.GC.1012.NGR.050'), ('STF.GC.1018.ADH.001', 'STF.GC.1018.ADH.001'),
             ('STF.GC.1021.CHK.001', 'STF.GC.1021.CHK.001'), ('STF.GC.1021.CHK.002', 'STF.GC.1021.CHK.002'),
             ('STF.GC.1021.KND.050', 'STF.GC.1021.KND.050'), ('STF.GC.1023.GGN.051', 'STF.GC.1023.GGN.051'),
             ('STF.GC.1024.KND.037', 'STF.GC.1024.KND.037'), ('STF.GC.1031.SLP.034', 'STF.GC.1031.SLP.034'),
             ('STF.GC.1031.THN.034', 'STF.GC.1031.THN.034'), ('STF.GC.1031.WRL.034', 'STF.GC.1031.WRL.034'),
             ('STF.GC.1032.AMD.050', 'STF.GC.1032.AMD.050'), ('STF.GC.1032.BLR.050', 'STF.GC.1032.BLR.050'),
             ('STF.GC.1032.BPL.050', 'STF.GC.1032.BPL.050'), ('STF.GC.1032.BPL.052', 'STF.GC.1032.BPL.052'),
             ('STF.GC.1032.CCN.050', 'STF.GC.1032.CCN.050'), ('STF.GC.1032.CCN.052', 'STF.GC.1032.CCN.052'),
             ('STF.GC.1032.CHN.050', 'STF.GC.1032.CHN.050'), ('STF.GC.1032.CHN.051', 'STF.GC.1032.CHN.051'),
             ('STF.GC.1032.CHN.052', 'STF.GC.1032.CHN.052'), ('STF.GC.1032.CHN.053 ', 'STF.GC.1032.CHN.053 '),
             ('STF.GC.1032.JPR.050', 'STF.GC.1032.JPR.050'), ('STF.GC.1032.JPR.052', 'STF.GC.1032.JPR.052'),
             ('STF.GC.1032.KHM.052', 'STF.GC.1032.KHM.052'), ('STF.GC.1032.KND.050', 'STF.GC.1032.KND.050'),
             ('STF.GC.1032.KND.051', 'STF.GC.1032.KND.051'), ('STF.GC.1032.KND.052', 'STF.GC.1032.KND.052'),
             ('STF.GC.1032.LKW.050', 'STF.GC.1032.LKW.050'), ('STF.GC.1032.NSK.051', 'TF.GC.1032.NSK.051'),
             ('STF.GC.1032.PTN.052', 'STF.GC.1032.PTN.052'), ('STF.GC.1032.VJW.050', 'STF.GC.1032.VJW.050'),
             ('STF.GC.1032.VJW.051', 'STF.GC.1032.VJW.051'), ('STF.GC.1032.VJW.052', 'STF.GC.1032.VJW.052'),
             ('STF.GC.1033.AMD.050', 'STF.GC.1033.AMD.050'), ('STF.GC.1033.AMD.055', 'STF.GC.1033.AMD.055'),
             ('STF.GC.1033.AMD.056', 'STF.GC.1033.AMD.056'), ('STF.GC.1033.AMD.058', 'STF.GC.1033.AMD.058'),
             ('STF.GC.1033.BPL.050', 'STF.GC.1033.BPL.050'), ('STF.GC.1033.BPL.054', 'STF.GC.1033.BPL.054'),
             ('STF.GC.1033.BPL.055', 'STF.GC.1033.BPL.055'), ('STF.GC.1033.BPL.056', 'STF.GC.1033.BPL.056'),
             ('STF.GC.1033.BPL.057', 'STF.GC.1033.BPL.057'), ('STF.GC.1033.BPL.058', 'STF.GC.1033.BPL.058'),
             ('STF.GC.1033.CHD.050', 'STF.GC.1033.CHD.050'), ('STF.GC.1033.CHD.055', 'STF.GC.1033.CHD.055'),
             ('STF.GC.1033.CHD.056', 'STF.GC.1033.CHD.056'), ('STF.GC.1033.CHN.001', 'STF.GC.1033.CHN.001'),
             ('STF.GC.1033.DEL.050', 'STF.GC.1033.DEL.050'), ('STF.GC.1033.DEL.054', 'STF.GC.1033.DEL.054'),
             ('STF.GC.1033.DEL.055', 'STF.GC.1033.DEL.055'), ('STF.GC.1033.DEL.056', 'STF.GC.1033.DEL.056'),
             ('STF.GC.1033.JMU.050', 'STF.GC.1033.JMU.050'), ('STF.GC.1033.JMU.056', 'STF.GC.1033.JMU.056'),
             ('STF.GC.1033.JPR.050', 'STF.GC.1033.JPR.050'), ('STF.GC.1033.JPR.055', 'STF.GC.1033.JPR.055'),
             ('STF.GC.1033.JPR.056', 'STF.GC.1033.JPR.056'), ('STF.GC.1033.KND.001', 'STF.GC.1033.KND.001'),
             ('STF.GC.1033.KND.050', 'STF.GC.1033.KND.050'), ('STF.GC.1033.KND.053', 'STF.GC.1033.KND.053'),
             ('STF.GC.1033.KND.055', 'STF.GC.1033.KND.055'), ('STF.GC.1033.KND.056', 'STF.GC.1033.KND.056'),
             ('STF.GC.1033.KND.058', 'STF.GC.1033.KND.058'), ('STF.GC.1033.KNL.050', 'STF.GC.1033.KNL.050'),
             ('STF.GC.1033.KNL.056', 'STF.GC.1033.KNL.056'), ('STF.GC.1033.LDH.050', 'STF.GC.1033.LDH.050'),
             ('STF.GC.1033.LDH.056', 'STF.GC.1033.LDH.056'), ('STF.GC.1033.LKW.050', 'STF.GC.1033.LKW.050'),
             ('STF.GC.1033.RPR.050', 'STF.GC.1033.RPR.050'), ('STF.GC.1033.RPR.055', 'STF.GC.1033.RPR.055'),
             ('STF.GC.1033.SML.050', 'STF.GC.1033.SML.050'), ('STF.GC.1033.SML.056', 'STF.GC.1033.SML.056'),
             ('STF.GC.1037.BLR.001', 'STF.GC.1037.BLR.001'), ('STF.GC.1044.BLR.001', 'STF.GC.1044.BLR.001'),
             ('STF.GC.1047.GRN.001', 'STF.GC.1047.GRN.001'), ('STF.GC.1049.HYD.052', 'STF.GC.1049.HYD.052'),
             ('STF.GC.1049.KND', 'STF.GC.1049.KND.050'), ('STF.GC.1049.KND.050', 'STF.GC.1049.KND.050'),
             ('STF.GC.1049.KND.051', 'STF.GC.1049.KND.051'), ('STF.GC.1049.KND.052', 'STF.GC.1049.KND.052'),
             ('STF.GC.1049.KND.053', 'STF.GC.1049.KND.053'), ('STF.GC.1049.KND.054', 'STF.GC.1049.KND.054'),
             ('STF.GC.1049.KSH.052', 'STF.GC.1049.KSH.052'), ('STF.GC.1049.PTN.052', 'STF.GC.1049.PTN.052'),
             ('STF.GC.1049.RPR.052', 'STF.GC.1049.RPR.052'), ('STF.GC.1049.UPR.052', 'STF.GC.1049.UPR.052'),
             ('STF.GC.1055.DEL.001', 'STF.GC.1055.DEL.001'), ('STF.GC.1055.KND.002 ', 'STF.GC.1055.KND.002 '),
             ('STF.GC.1055.MAH.001', 'STF.GC.1055.MAH.001'), ('STF.GC.1055.TLN.001', 'STF.GC.1055.TLN.001'),
             ('STF.GC.1057.KND.001', 'STF.GC.1057.KND.001'), ('STF.GC.1058.KND.048', 'STF.GC.1058.KND.048'),
             ('STF.GC.1065.KND.050', 'STF.GC.1065.KND.050'), ('STF.GC.1069.THN.002', 'STF.GC.1069.THN.002'),
             ('STF.GC.5004.MIT.001', 'STF.GC.5004.MIT.001'), ('STF.GC.9025.PNE.001', 'STF.GC.9025.PNE.001'),
             ('STF.GC.MEML.KND.001', 'STF.GC.MEML.KND.001'), ('STF.MM.AUTO.BLR.059', 'STF.MM.AUTO.BLR.059'),
             ('STF.MM.AUTO.CHN.001', 'STF.MM.AUTO.CHN.001'), ('STF.MM.AUTO.CHN.060', 'STF.MM.AUTO.CHN.060'),
             ('STF.MM.AUTO.HRD.003', 'STF.MM.AUTO.HRD.003'), ('STF.MM.AUTO.HRD.060', 'STF.MM.AUTO.HRD.060'),
             ('STF.MM.AUTO.IGT.002', 'STF.MM.AUTO.IGT.002'), ('STF.MM.AUTO.KND.001', 'STF.MM.AUTO.KND.001'),
             ('STF.MM.AUTO.KND.004', 'STF.MM.AUTO.KND.004'), ('STF.MM.AUTO.KND.005', 'STF.MM.AUTO.KND.005'),
             ('STF.MM.AUTO.KND.006 ', 'STF.MM.AUTO.KND.006 '), ('STF.MM.AUTO.KND.007', 'STF.MM.AUTO.KND.007'),
             ('STF.MM.AUTO.KND.046', 'STF.MM.AUTO.KND.046'), ('STF.MM.AUTO.KND.047', 'STF.MM.AUTO.KND.047'),
             ('STF.MM.AUTO.KND.048', 'STF.MM.AUTO.KND.048'), ('STF.MM.AUTO.KND.049', 'STF.MM.AUTO.KND.049'),
             ('STF.MM.AUTO.KND.050', 'STF.MM.AUTO.KND.050'), ('STF.MM.AUTO.KND.051', 'STF.MM.AUTO.KND.051'),
             ('STF.MM.AUTO.KND.052', 'STF.MM.AUTO.KND.052'), ('STF.MM.AUTO.KND.053', 'STF.MM.AUTO.KND.053'),
             ('STF.MM.AUTO.KND.054', 'STF.MM.AUTO.KND.054'), ('STF.MM.AUTO.KND.055', 'STF.MM.AUTO.KND.055'),
             ('STF.MM.AUTO.KND.057', 'STF.MM.AUTO.KND.057'), ('STF.MM.AUTO.KND.058', 'STF.MM.AUTO.KND.058'),
             ('STF.MM.AUTO.KND.059', 'STF.MM.AUTO.KND.059'), ('STF.MM.AUTO.KND.060', 'STF.MM.AUTO.KND.060'),
             ('STF.MM.AUTO.KND.061', 'STF.MM.AUTO.KND.061'), ('STF.MM.AUTO.KND.062', 'STF.MM.AUTO.KND.062'),
             ('STF.MM.AUTO.KND.063', 'STF.MM.AUTO.KND.063'), ('STF.MM.AUTO.KND.064', 'STF.MM.AUTO.KND.064'),
             ('STF.MM.AUTO.KND.065', 'STF.MM.AUTO.KND.065'), ('STF.MM.AUTO.KND.066', 'STF.MM.AUTO.KND.066'),
             ('STF.MM.AUTO.KOL.060', 'STF.MM.AUTO.KOL.060'), ('STF.MM.AUTO.LKW.060', 'STF.MM.AUTO.LKW.060'),
             ('STF.MM.AUTO.MTBD', 'STF.MM.AUTO.MTBD'), ('STF.MM.AUTO.NSK.001', 'STF.MM.AUTO.NSK.001'),
             ('STF.MM.AUTO.NSK.049', 'STF.MM.AUTO.NSK.049'), ('STF.MM.AUTO.NSK.054', 'STF.MM.AUTO.NSK.054'),
             ('STF.MM.AUTO.PNE.053', 'STF.MM.AUTO.PNE.053'), ('STF.MM.AUTO.PNE.054', 'STF.MM.AUTO.PNE.054'),
             ('STF.MM.AUTO.WRL.001', 'STF.MM.AUTO.WRL.001'), ('STF.MM.CORP.CHN.062', 'STF.MM.CORP.CHN.062'),
             ('STF.MM.CORP.FRD.064', 'STF.MM.CORP.FRD.064'), ('STF.MM.CORP.KND.005', 'STF.MM.CORP.KND.005'),
             ('STF.MM.CORP.KND.031', 'STF.MM.CORP.KND.031'), ('STF.MM.CORP.KND.033', 'STF.MM.CORP.KND.033'),
             ('STF.MM.CORP.KND.051', 'STF.MM.CORP.KND.051'), ('STF.MM.CORP.KND.052', 'STF.MM.CORP.KND.052'),
             ('STF.MM.CORP.KND.053', 'STF.MM.CORP.KND.053'), ('STF.MM.CORP.KND.054', 'STF.MM.CORP.KND.054'),
             ('STF.MM.CORP.KND.055', 'STF.MM.CORP.KND.055'), ('STF.MM.CORP.KND.056', 'STF.MM.CORP.KND.056'),
             ('STF.MM.CORP.KND.057', 'STF.MM.CORP.KND.057'), ('STF.MM.CORP.KND.058', 'STF.MM.CORP.KND.058'),
             ('STF.MM.CORP.KND.059', 'STF.MM.CORP.KND.059'), ('STF.MM.CORP.KND.060', 'STF.MM.CORP.KND.060'),
             ('STF.MM.CORP.KND.061', 'STF.MM.CORP.KND.061'), ('STF.MM.CORP.KND.062', 'STF.MM.CORP.KND.062'),
             ('STF.MM.CORP.KND.063', 'STF.MM.CORP.KND.063'), ('STF.MM.CORP.KND.064', 'STF.MM.CORP.KND.064'),
             ('STF.MM.CORP.KND.065', 'STF.MM.CORP.KND.065'), ('STF.MM.CORP.KND.066', 'STF.MM.CORP.KND.066'),
             ('STF.MM.CORP.KND.070', 'STF.MM.CORP.KND.070'), ('STF.MM.CORP.MIT.069', 'STF.MM.CORP.MIT.069'),
             ('STF.MM.CORP.PNE.063', 'STF.MM.CORP.PNE.063'), ('STF.MM.CORP.SWR.063', 'STF.MM.CORP.SWR.063'),
             ('STF.MM.CORP.SWR.064', 'STF.MM.CORP.SWR.064'), ('STF.MM.CORP.SWR.069', 'STF.MM.CORP.SWR.069'),
             ('STF.MM.CORP.SWR.070', 'STF.MM.CORP.SWR.070'), ('STF.MM.CORP.WRL.004', 'STF.MM.CORP.WRL.004'),
             ('STF.MM.CORP.WRL.060', 'STF.MM.CORP.WRL.060'), ('STF.MM.FESD.APR.001', 'STF.MM.FESD.APR.001'),
             ('STF.MM.FESD.APR.003', 'STF.MM.FESD.APR.003'), ('STF.MM.FESD.APR.004', 'STF.MM.FESD.APR.004'),
             ('STF.MM.FESD.APR.005', 'STF.MM.FESD.APR.005'), ('STF.MM.FESD.ASM.001', 'STF.MM.FESD.ASM.001'),
             ('STF.MM.FESD.ASM.006', 'STF.MM.FESD.ASM.006'), ('STF.MM.FESD.BHR.001', 'STF.MM.FESD.BHR.001'),
             ('STF.MM.FESD.BHR.003 ', 'STF.MM.FESD.BHR.003 '), ('STF.MM.FESD.BHR.003\xa0', 'STF.MM.FESD.BHR.003\xa0'),
             ('STF.MM.FESD.BHR.004', 'STF.MM.FESD.BHR.004'), ('STF.MM.FESD.BHR.005', 'STF.MM.FESD.BHR.005'),
             ('STF.MM.FESD.BHR.006', 'STF.MM.FESD.BHR.006'), ('STF.MM.FESD.CTG.001', 'STF.MM.FESD.CTG.001'),
             ('STF.MM.FESD.CTG.003 ', 'STF.MM.FESD.CTG.003 '), ('STF.MM.FESD.CTG.003\xa0', 'STF.MM.FESD.CTG.003\xa0'),
             ('STF.MM.FESD.CTG.004', 'STF.MM.FESD.CTG.004'), ('STF.MM.FESD.CTG.005', 'STF.MM.FESD.CTG.005'),
             ('STF.MM.FESD.CTG.006', 'STF.MM.FESD.CTG.006'), ('STF.MM.FESD.DEL.001', 'STF.MM.FESD.DEL.001'),
             ('STF.MM.FESD.DEL.004', 'STF.MM.FESD.DEL.004'), ('STF.MM.FESD.DEL.005', 'STF.MM.FESD.DEL.005'),
             ('STF.MM.FESD.GUJ.001', 'STF.MM.FESD.GUJ.001'), ('STF.MM.FESD.GUJ.004', 'STF.MM.FESD.GUJ.004'),
             ('STF.MM.FESD.GUJ.005', 'STF.MM.FESD.GUJ.005'), ('STF.MM.FESD.GUJ.006', 'STF.MM.FESD.GUJ.006'),
             ('STF.MM.FESD.HAR.001', 'STF.MM.FESD.HAR.001'), ('STF.MM.FESD.HAR.003', 'STF.MM.FESD.HAR.003'),
             ('STF.MM.FESD.HAR.004', 'STF.MM.FESD.HAR.004'), ('STF.MM.FESD.HAR.005', 'STF.MM.FESD.HAR.005'),
             ('STF.MM.FESD.HAR.006', 'STF.MM.FESD.HAR.006'), ('STF.MM.FESD.JHK.001', 'STF.MM.FESD.JHK.001'),
             ('STF.MM.FESD.JHK.004', 'STF.MM.FESD.JHK.004'), ('STF.MM.FESD.JHK.005', 'STF.MM.FESD.JHK.005'),
             ('STF.MM.FESD.KAR.001', 'STF.MM.FESD.KAR.001'), ('STF.MM.FESD.KAR.004', 'STF.MM.FESD.KAR.004'),
             ('STF.MM.FESD.KAR.005', 'STF.MM.FESD.KAR.005'), ('STF.MM.FESD.KAR.006', 'STF.MM.FESD.KAR.006'),
             ('STF.MM.FESD.KND', 'STF.MM.FESD.KND.002'), ('STF.MM.FESD.KND.002', 'STF.MM.FESD.KND.002'),
             ('STF.MM.FESD.KND.008', 'STF.MM.FESD.KND.008'), ('STF.MM.FESD.KND.009', 'STF.MM.FESD.KND.009'),
             ('STF.MM.FESD.KND.017', 'STF.MM.FESD.KND.017'), ('STF.MM.FESD.KND.048', 'STF.MM.FESD.KND.048'),
             ('STF.MM.FESD.KND.050', 'STF.MM.FESD.KND.050'), ('STF.MM.FESD.KND.051', 'STF.MM.FESD.KND.051'),
             ('STF.MM.FESD.KND.052', 'STF.MM.FESD.KND.052'), ('STF.MM.FESD.KND.053', 'STF.MM.FESD.KND.053'),
             ('STF.MM.FESD.KND.054', 'STF.MM.FESD.KND.054'), ('STF.MM.FESD.KND.055', 'STF.MM.FESD.KND.055'),
             ('STF.MM.FESD.KND.056', 'STF.MM.FESD.KND.056'), ('STF.MM.FESD.KNH.001', 'STF.MM.FESD.KNH.001'),
             ('STF.MM.FESD.KOL.051', 'STF.MM.FESD.KOL.051'), ('STF.MM.FESD.LKW.051', 'STF.MM.FESD.LKW.051'),
             ('STF.MM.FESD.MAH.001', 'STF.MM.FESD.MAH.001'), ('STF.MM.FESD.MAH.003', 'STF.MM.FESD.MAH.003'),
             ('STF.MM.FESD.MAH.004', 'STF.MM.FESD.MAH.004'), ('STF.MM.FESD.MAH.005', 'STF.MM.FESD.MAH.005'),
             ('STF.MM.FESD.MAH.006', 'STF.MM.FESD.MAH.006'), ('STF.MM.FESD.MAH.007', 'STF.MM.FESD.MAH.007'),
             ('STF.MM.FESD.MPR.001', 'STF.MM.FESD.MPR.001'), ('STF.MM.FESD.MPR.003 ', 'STF.MM.FESD.MPR.003 '),
             ('STF.MM.FESD.MPR.003\xa0', 'STF.MM.FESD.MPR.003\xa0'), ('STF.MM.FESD.MPR.004', 'STF.MM.FESD.MPR.004'),
             ('STF.MM.FESD.MPR.005', 'STF.MM.FESD.MPR.005'), ('STF.MM.FESD.MPR.006', 'STF.MM.FESD.MPR.006'),
             ('STF.MM.FESD.NGP.001', 'STF.MM.FESD.NGP.001'), ('STF.MM.FESD.ORS.001', 'STF.MM.FESD.ORS.001'),
             ('STF.MM.FESD.ORS.003 ', 'STF.MM.FESD.ORS.003 '), ('STF.MM.FESD.ORS.003\xa0', 'STF.MM.FESD.ORS.003\xa0'),
             ('STF.MM.FESD.ORS.004', 'STF.MM.FESD.ORS.004'), ('STF.MM.FESD.ORS.005', 'STF.MM.FESD.ORS.005'),
             ('STF.MM.FESD.ORS.006', 'STF.MM.FESD.ORS.006'), ('STF.MM.FESD.PJB.001', 'STF.MM.FESD.PJB.001'),
             ('STF.MM.FESD.PJB.003', 'STF.MM.FESD.PJB.003 '), ('STF.MM.FESD.PJB.003\xa0', 'STF.MM.FESD.PJB.003\xa0'),
             ('STF.MM.FESD.PJB.004', 'STF.MM.FESD.PJB.004'), ('STF.MM.FESD.PJB.005', 'STF.MM.FESD.PJB.005'),
             ('STF.MM.FESD.PJB.006', 'STF.MM.FESD.PJB.006'), ('STF.MM.FESD.PJB.007', 'STF.MM.FESD.PJB.007'),
             ('STF.MM.FESD.PNE.001', 'STF.MM.FESD.PNE.001'), ('STF.MM.FESD.RAJ.001', 'STF.MM.FESD.RAJ.001'),
             ('STF.MM.FESD.RAJ.004', 'STF.MM.FESD.RAJ.004'), ('STF.MM.FESD.RAJ.005', 'STF.MM.FESD.RAJ.005'),
             ('STF.MM.FESD.RAJ.006', 'STF.MM.FESD.RAJ.006'), ('STF.MM.FESD.TLG.001', 'STF.MM.FESD.TLG.001'),
             ('STF.MM.FESD.TLG.003', 'STF.MM.FESD.TLG.003'), ('STF.MM.FESD.TLG.004', 'STF.MM.FESD.TLG.004'),
             ('STF.MM.FESD.TLG.005', 'STF.MM.FESD.TLG.005'), ('STF.MM.FESD.TLG.006', 'STF.MM.FESD.TLG.006'),
             ('STF.MM.FESD.TLN.001', 'STF.MM.FESD.TLN.001'), ('STF.MM.FESD.TLN.003', 'STF.MM.FESD.TLN.003'),
             ('STF.MM.FESD.TLN.004', 'STF.MM.FESD.TLN.004'), ('STF.MM.FESD.TLN.005', 'STF.MM.FESD.TLN.005'),
             ('STF.MM.FESD.TLN.006', 'STF.MM.FESD.TLN.006'), ('STF.MM.FESD.UPR.001', 'STF.MM.FESD.UPR.001'),
             ('STF.MM.FESD.UPR.003', 'STF.MM.FESD.UPR.003'), ('STF.MM.FESD.UPR.004', 'STF.MM.FESD.UPR.004'),
             ('STF.MM.FESD.UPR.005', 'STF.MM.FESD.UPR.005'), ('STF.MM.FESD.UPR.006', 'STF.MM.FESD.UPR.006'),
             ('STF.MM.FESD.WBL.001', 'STF.MM.FESD.WBL.001'), ('STF.MM.FESD.WBL.003', 'STF.MM.FESD.WBL.003'),
             ('STF.MM.FESD.WBL.004', 'STF.MM.FESD.WBL.004'), ('STF.MM.FESD.WBL.005', 'STF.MM.FESD.WBL.005'),
             ('STF.MM.FESD.WBL.006', 'STF.MM.FESD.WBL.006'), ('STF.MM.MTWD.PNE.001', 'STF.MM.MTWD.PNE.001'),
             ('STF.MM.MTWD.PNE.002', 'STF.MM.MTWD.PNE.002'), ('STF.MM.SBUD.KND.045', 'STF.MM.SBUD.KND.045'),
             ('STF.MM.SBUD.KND.054', 'STF.MM.SBUD.KND.054'), ('STF.MSM.FESD.MAH.004', 'STF.MM.FESD.MAH.004')])
        self.chtext69 = OrderedDict(
            [('00', 'Default'), ('01', 'Kandivali'), ('02', 'AS-Kandivali'), ('03', 'Bangalore'), ('04', 'Chakan'),
             ('05', 'Chennai'), ('06', 'FES - Kandivali'), ('07', 'Haridwar'), ('08', 'Igatpuri'), ('09', 'Jaipur'),
             ('10', 'Kanhe'), ('11', 'MADPL - Kandivali'), ('12', 'MRV - Kandivali'), ('13', 'Nagpur'),
             ('14', 'Nashik'), ('15', 'Rudrapur'), ('16', 'Sewree'), ('17', 'Wadgaon'), ('18', 'Worli'),
             ('19', 'Zaheerabad'), ('20', 'Kandivali1'), ('21', 'Sewree1'), ('22', 'Delhi'), ('23', 'Vashi'),
             ('24', 'Bhiwandi'), ('25', 'Thane- MIBL'), ('26', 'Purnia'), ('27', 'Raipur'), ('28', 'New Delhi'),
             ('29', 'Ahmedabad'), ('30', 'Hissar'), ('31', 'Mandi'), ('32', 'Jammu & Kashmir'), ('33', 'Ranchi'),
             ('34', 'Bangalore'), ('35', 'Bellary'), ('36', 'Cochin'), ('37', 'Gwalior'), ('38', 'Indore'),
             ('39', 'Akarudi'), ('40', 'Akola'), ('41', 'Aurangabad'), ('42', 'Bhiwandi'), ('43', 'Kanhe'),
             ('44', 'Kolhapur'), ('45', 'Nanded'), ('46', 'Pune'), ('47', 'Vadgaon'), ('48', 'Bhubaneswar'),
             ('49', 'Bhatinda'), ('50', 'Jalandhar'), ('51', 'Ganganagar'), ('52', 'Jaipur'), ('53', 'Kota'),
             ('54', 'Madurai'), ('55', 'Salem'), ('56', 'Hyderabad'), ('57', 'Agartala'), ('58', 'Gorakhpur'),
             ('59', 'Kanpur'), ('60', 'Lucknow'), ('61', 'Haldwani'), ('62', 'Dehradun'), ('63', 'Asansol'),
             ('64', 'Vijayawada'), ('65', 'Vizag'), ('66', 'Guwahati'), ('67', 'Silchar'), ('68', 'Patna'),
             ('69', 'Cuttack'), ('70  ', 'Goregaon'), ('71', 'Solapur'), ('72', 'Bhopal'), ('73', 'Rajkot'),
             ('74', 'Thane'), ('75', 'KARNATAKA'), ('76', 'TRICHY'), ('77', 'Mohali'), ('78', 'Jabalpur'),
             ('79', 'Baroda'), ('80', 'Surat'), ('81', 'Coimbatore'), ('82', 'Bikaner')])
        self.chtext70 = OrderedDict(
            [('01', 'Kandivali- HO'), ('02', 'AS-Kandivali'), ('020', 'Thane'), ('03', 'Nagpur'), ('04', 'Nashik'),
             ('05', 'Zaheerabad'), ('06', 'Chakan'), ('07', 'Haridwar'), ('08', 'FES-Kandivali'), ('09', 'Rudrapur'),
             ('10', 'Chennai'), ('100', 'THARAD'), ('1000', 'Lasalgaon'), ('1001', 'Phulambri'), ('1002', 'Selu'),
             ('1003', 'Shevgaon'), ('1004', 'Palam'), ('1005', 'Soegaon'), ('1006', 'Sonpeth'), ('1007', 'Omerga'),
             ('1008', 'Vaijapur'), ('1009', 'Partur'), ('101', 'BEHRUCH(NEW)'), ('1010', 'Tuljapur'),
             ('1011', 'Mangrulpir'), ('1012', 'Badnapur'), ('1013', 'Georai'), ('1014', 'Dharmabad'),
             ('1015', 'Bhokardan'), ('1016', 'Paranda'), ('1017', 'Naigaon'), ('1018', 'Kalamb'), ('1019', 'Daryapur'),
             ('102', 'BHOPAL'), ('1020', 'Umarkhed'), ('1021', 'Kelapur'), ('1022', 'Sangrampur'), ('1023', 'Tiroda'),
             ('1024', 'Ner'), ('1025', 'Bhandra'), ('1026', 'Ralegaon'), ('1027', 'Kinwat'), ('1028', 'Osmanabad'),
             ('1029', 'Dhamangaon'), ('103', 'HABIBGUNJ'), ('1030', 'Khultabad'), ('1031', 'Shivna'), ('1032', 'Mahur'),
             ('1033', 'Mukhed'), ('1034', 'Bidkin'), ('1035', 'Himayatngar'), ('1036', 'Thirthpuri'),
             ('1037', 'Omarga'), ('1038', 'Barshitakli'), ('1039', 'Alwar'), ('104', 'RAU'), ('1040', 'Pirawa'),
             ('1041', 'Janagam'), ('1042', 'Suryapet'), ('1043', 'Gadchandur\xa0'), ('1044', 'Thanjavur'),
             ('1045', 'Durgapur'), ('105', 'HABIBGUNG'), ('106', 'DEPALPUR'), ('107', 'DEWAS'), ('108', 'Vidisha'),
             ('109', 'DABRA'), ('11', 'Igatpuri'), ('110', 'KHANDWA'), ('111', 'Gwalior'), ('112', 'MANDSOUR'),
             ('113', 'SHIVPURI'), ('114', 'MORENA'), ('115', 'GADARWARA'), ('116', 'REWA'), ('117', 'CHHATARPUR'),
             ('118', 'SHAJAPUR'), ('119', 'CHAKGHAT'), ('12', 'Bengaluru'), ('120', 'REWA'), ('121', 'UJJAIN'),
             ('122', 'TPT NAGAR INDORE'), ('123', 'PALAMPUR'), ('124', 'BADDI'), ('125', 'RAMPUR'), ('126', 'CHAMBA'),
             ('127', 'SHIMLA RO'), ('128', 'KULLU'), ('129', 'NALAGARH (NEW)'), ('13', 'Andheri'),
             ('130', 'SIMLA (THOEG)'), ('131', 'PATIALA'), ('132', 'BARNALA'), ('133', 'NAWANSHAHR'),
             ('134', 'RAJPURA'), ('135', 'ZIRAKPUR'), ('136', 'FARIDKOT'), ('137', 'LUDHIANA - CDG'),
             ('138', 'BARAMATI'), ('139', 'SATANA'), ('14', 'Chakan-Kanhe'), ('140', 'SHIRPUR'), ('141', 'SINNAR'),
             ('142', 'SWARGATE'), ('143', 'KOLHAPUR'), ('144', 'WASHIM'), ('145', 'PIMPRI'), ('146', 'SOLAPUR'),
             ('147', 'Pune'), ('148', 'SANGLI'), ('149', 'SHIKRAPUR'), ('15', 'FES-Jaipur'), ('150', 'AHMEDNAGAR'),
             ('151', 'SATARA'), ('152', 'JATH'), ('153', 'Jabalpur'), ('154', 'Barshi'), ('155', 'BHILAI'),
             ('156', 'UDHAMPUR'), ('157', 'SIDHI'), ('158', 'Manglore'), ('159', 'Bundi'), ('16', 'MTWL -Chkan'),
             ('160', 'kotputli'), ('161', 'Mehkar ( buldana )'), ('162', 'Barelli'), ('163', 'BALRAMPUR'),
             ('164', 'BALODABAZAR'), ('165', 'Balaod'), ('166', 'CHHATTISGARH'), ('167', 'DANTEWADA'),
             ('168', 'BHATAPARA'), ('169', 'CHHATTISGARH'), ('17', 'MTBL-Pune'), ('170', 'PAKHANJORE'),
             ('171', 'PENDRA'), ('172', 'ABOHAR'), ('173', 'Rajkot'), ('174', 'VIRANIA RCC'), ('175', 'BAYAD RCC'),
             ('176', 'BODELI'), ('177', 'RADHANPUR'), ('178', 'Palanpur'), ('179', 'IDAR'), ('18', 'DELHI'),
             ('180', 'Surat'), ('181', 'ANKLESHWAR'), ('182', 'Dhanera'), ('183', 'Godhra'), ('184', 'Anand'),
             ('185', 'VANSDA'), ('186', 'Vyara'), ('187', 'BARODA'), ('188', 'GURGAON'), ('189', 'Jind'),
             ('19', 'Worli'), ('190', 'Ambala'), ('191', 'Kaithal'), ('192', 'SIRSA - DLH'), ('193', 'JASOLA'),
             ('194', 'DELHI BAHADURGAR'), ('195', 'Bilaspur'), ('196', 'PONTA SAHIB'), ('197', 'RECKONG PEO'),
             ('198', 'UNA SB'), ('199', 'Srinagar'), ('20', 'Kandivali-Emply.Benefit'), ('200', 'ALIRAJPUR'),
             ('201', 'JUNNARDEO'), ('202', 'Mauganj'), ('203', 'DHAMNOD'), ('204', 'KHILCHIPUR'), ('205', 'UJJAIN'),
             ('206', 'SHUJALPUR'), ('207', 'Khargone'), ('208', 'HARDA'), ('209', 'SEHORE'), ('21', 'Vijayawada'),
             ('210', 'Sagar'), ('211', 'KUSHMODHA'), ('212', 'Ashoknagar'), ('213', 'SENDHWA (MP)'), ('214', 'Waidhan'),
             ('215', 'GUNA-MP'), ('216', 'BHANPURA'), ('217', 'Barwaha'), ('218', 'KHATEGAON'), ('219', 'Jabalpur'),
             ('22', 'Vizag'), ('220', 'SATNA'), ('221', 'DATIA'), ('222', 'RATLAM-MP'), ('223', 'Waidhan'),
             ('224', 'UMARIA'), ('225', 'SEONI'), ('226', 'SASWAD'), ('227', 'GHATKOPAR'), ('228', 'BORIVALI'),
             ('229', 'Aurangabad'), ('23', 'Guwahati'), ('230', 'Panvel'), ('231', 'Khopoli'), ('232', 'SUJANGARH'),
             ('233', 'Banswara'), ('234', 'AJMER'), ('235', 'CHHABRA'), ('236', 'JHALAWAR'), ('237', 'DHOLPUR'),
             ('238', 'CHOMU'), ('239', 'RAJSAMAND'), ('24', 'Silchar'), ('240', 'ABU ROAD'), ('241', 'Baran'),
             ('242', 'BEHROR'), ('243', 'BHIWADI'), ('244', 'JHUNJHUNU'), ('245', 'KUCHAMAN CITY'),
             ('246', 'JALORE RCC'), ('247', 'JAISALMER'), ('248', 'Phalodi'), ('249', 'Jodhpur'), ('25', 'Patna'),
             ('250', 'Dungargarh'), ('251', 'SIKAR'), ('252', 'HANUMANGARH'), ('253', 'Uttar Pradesh'),
             ('254', 'Sahibabad'), ('255', 'CHHATTISGARH'), ('256', 'GUJARAT'), ('257', 'DATIA'), ('258', 'Ghaziabad'),
             ('259', 'Faridabad'), ('26', 'Purnia'), ('260', 'Salarpur'), ('261', 'Jagdalpur'), ('262', 'Chandkheda'),
             ('263', 'JODHPUR'), ('264', 'Kurukshetra'), ('265', 'Ankleshwar'), ('266', 'SHIVPURI -MP'),
             ('267', 'BERASIA'), ('268', 'GANJBASODA SB'), ('269', 'GOHAD'), ('27', 'Raipur'), ('270', 'Barwani'),
             ('271', 'NASRULLAHGANJ'), ('272', 'SABALGARH SB'), ('273', 'PIPARIYA'), ('274', 'BARMER'),
             ('275', 'Boisar'), ('276', 'jharsa'), ('277', 'BIKANER'), ('278', 'RAISEN'), ('279', 'MANDLA'),
             ('28', 'New Delhi'), ('280', 'BHIND'), ('281', 'HOSHANGABAD'), ('282', 'SEONDHA'), ('283', 'Sheopur'),
             ('284', 'Bad Malhera'), ('285', 'Mungeli'), ('29', 'Dehgam'), ('299', 'BELAPUR'), ('30', 'Hissar'),
             ('300', 'Kushtagi'), ('301', 'Holenarasipura'), ('302', 'Davangere (Karnataka)'), ('303', 'Mugur'),
             ('304', 'Aland'), ('305', 'Karad'), ('306', 'Malgaon'), ('308', 'BHANDARA'), ('309', 'COIMBATORE'),
             ('31', 'Mandi'), ('310', 'CHINDWARA'), ('311', 'GURGAON'), ('312', 'DHANERA'), ('313', 'Sonipat'),
             ('314', 'BARODA'), ('315', 'MODASA'), ('316', 'HAZIRA'), ('317', 'Katni'), ('318', 'Seoni'),
             ('319', 'Chikhali'), ('32', 'SHIMLA BRANCH'), ('320', 'Pratapgarh'), ('321', 'KUCHAMAN CITY'),
             ('322', 'Baran'), ('323', 'KHATEGAON'), ('324', 'Dehgam'), ('325', 'CHOMU'), ('326', 'ZALOD'),
             ('327', 'KARHAL'), ('328', 'VASTRAL'), ('329', 'Bhilwara two'), ('33', 'Ranchi'), ('330', 'Jodhpur two'),
             ('331', 'HAPUR'), ('332', 'Arni'), ('333', 'Arvi'), ('334', 'Bramhpuri'), ('335', 'Chandrapur'),
             ('336', 'Chamorshi'), ('337', 'Deori'), ('338', 'Darwha'), ('339', 'Gadchiroli'), ('34', 'Bengaluru'),
             ('340', 'Gondpipri'), ('341', 'Gondia'), ('342', 'Hinganghat'), ('343', 'Katol'), ('344', 'Lakhandur'),
             ('345', 'Narkhed'), ('346', 'Ramtek'), ('347', 'Saoner'), ('348', 'Samudrpur'), ('349', 'Tumsar'),
             ('35', 'Ballari'), ('350', 'Umred'), ('351', 'Wani'), ('352', 'Wardha'), ('353', 'Warora'),
             ('354', 'Yavatmal'), ('355', 'Ahmedpur'), ('356', 'Akot'), ('357', 'Amalner'), ('358', 'Ambad'),
             ('359', 'Aundha'), ('36', 'Cochin'), ('360', 'Ausa'), ('361', 'Basmath'), ('362', 'Bhadgaon'),
             ('363', 'Chakur'), ('364', 'Chinchwad'), ('365', 'Deulgaon'), ('366', 'Dhule'), ('367', 'GANGAPUR'),
             ('368', 'Hadgaon'), ('369', 'Hingoli'), ('37', 'Gwalior'), ('370', 'Jalna'), ('371', 'Jintur'),
             ('372', 'Kalamnuri'), ('373', 'Kandhar'), ('374', 'Kannad'), ('375', 'Karanja Lad'), ('376', 'Khamgaon'),
             ('377', 'Lasur Station'), ('378', 'Latur'), ('379', 'Loha'), ('38', 'Indore'), ('380', 'Majalgaon'),
             ('381', 'Malkapur'), ('382', 'Mantha'), ('383', 'Morshi'), ('384', 'Nilanga'), ('385', 'Edappally'),
             ('386', 'Trichy'), ('387', 'Kerala'), ('388', 'gateway'), ('389', 'Sewri'), ('39', 'Akarudi'),
             ('390', 'Secunderabad'), ('391', 'Mohali'), ('392', 'Noida'), ('393', 'Varanasi'), ('394', 'Chandigarh'),
             ('395', 'Newasa'), ('396', 'Gadhinglaj'), ('397', 'Kupwara'), ('398', 'Budgam'), ('399', 'Kulgam'),
             ('40', 'Akola'), ('400', 'Dungargarh'), ('401', 'Akluj'), ('402', 'Ichalkaranji'), ('403', 'Narayangaon'),
             ('404', 'Shrigonda'), ('405', 'Kawardha'), ('406', 'pithora'), ('407', 'Jalore'), ('408', 'Sagwara'),
             ('409', 'Balesar'), ('41', 'Aurangabad'), ('410', 'Fatehnagar'), ('411', 'Jammu'), ('412', 'Banswara'),
             ('413', 'SULEPETH'), ('414', 'CHITTAPUR'), ('415', 'NELOGI'), ('416', 'KALBURGI'), ('417', 'ANDOLA'),
             ('418', 'KALGI'), ('419', 'MUDHOL'), ('42', 'Bhiwandi'), ('420', 'JAGALUR'), ('421', 'Hunasagi'),
             ('422', 'YADGIR'), ('423', 'FARHATABAD'), ('424', 'RAMANATHAPURA'), ('425', 'HITNAL'), ('426', 'MYSORE'),
             ('427', 'HUNSUR'), ('428', 'TERAKANAMIBI'), ('429', 'MYSORE'), ('43', 'Kanhe'), ('430', 'PATTAN'),
             ('431', 'MYSORE'), ('432', 'YADGIR'), ('433', 'RAISINGHNAGAR'), ('434', 'UDIGALA'), ('435', 'MYSORE'),
             ('436', 'MYSORE'), ('437', 'JOJI'), ('438', 'KARNATAKA'), ('439', 'Sangamner'), ('44', 'Kolhapur'),
             ('440', 'Junagadh'), ('441', 'Ojtoo'), ('442', 'reckong peo'), ('443', 'Bharuch'), ('444', 'GANDERBAL'),
             ('445', 'Sikar'), ('446', 'jhalawar'), ('447', 'SATNA'), ('448', 'phalodi'), ('449', 'BARUCH'),
             ('45', 'Nanded'), ('450', 'khatgora'), ('451', 'bhatgaon'), ('452', 'HASSAN'), ('453', 'RAICHUR (KAR)'),
             ('454', 'Gilesuguru'), ('455', 'GULBARGA'), ('456', 'Moka'), ('457', 'GANDHINAGAR'), ('458', 'Chhabra'),
             ('459', 'Reasi'), ('46', 'Pune'), ('460', 'Kishtwar'), ('461', 'Poonch Branch'), ('462', 'Barna'),
             ('463', 'Abu Road'), ('47', 'Vadgaon'), ('48', 'Bhubaneswar'), ('49', 'Bhatinda'), ('50', 'Jalandhar'),
             ('501', 'PRITAMPURA'), ('502', 'NANGLOI'), ('503', 'Alipur'), ('504', 'PALAM SB'), ('505', 'Delhi (RO)'),
             ('506', 'NAJAFGARH'), ('507', 'GOKULPURI'), ('508', 'Rudrapur'), ('509', 'MAHIPALPUR'),
             ('51', 'Ganganagar'), ('510', 'Punjabi Bagh'), ('511', 'MODEL TOWN'), ('52', 'Jaipur'), ('53', 'Kota'),
             ('533', 'Warora'), ('54', 'Madurai'), ('55', 'Salem'), ('56', 'Hyderabad'), ('57', 'Agartala'),
             ('58', 'Gorakhpur'), ('59', 'Kanpur'), ('60', 'Lucknow'), ('61', 'Haldwani'), ('62', 'Dehradun'),
             ('63', 'Asansol'), ('64', 'Cuttack'), ('65', 'Powai'), ('66', 'Amravati'), ('67', 'Goregaon'),
             ('68', 'KAMREJ'), ('69', 'VESU'), ('70', 'GANDHIDHAM'), ('71', 'MANINAGAR'), ('72', 'ROPAR'),
             ('73', 'PATHALGAON'), ('74', 'SARAIPALI'), ('75', 'RAIGARH'), ('76', 'MAHASAMUND'), ('77', 'KHAIRAGARH'),
             ('78', 'AMBIKAPUR'), ('79', 'BHIWANI'), ('80', 'Bilaspur'), ('800', 'Hampapura'), ('801', 'Karur'),
             ('802', 'Agar'), ('81', 'Chandigarh'), ('812', 'Hayyal'), ('813', 'Salegame'), ('814', 'Salligrama'),
             ('815', 'Hebbal'), ('816', 'Koregaon'), ('82', 'CHANDIGARH'), ('83', 'KANKER'), ('834', 'Nilanga'),
             ('835', 'Pachod'), ('836', 'Paithan'), ('837', 'Paratwada'), ('838', 'Pathari'), ('839', 'Patur'),
             ('84', 'DHAMTARI'), ('840', 'Pishor'), ('841', 'Purna'), ('842', 'Rajur'), ('843', 'Risod'),
             ('844', 'Sangamner'), ('845', 'Sengaon'), ('846', 'Shindkhedraja'), ('847', 'Shiruranantpal'),
             ('848', 'Shrirampur'), ('849', 'Sillod'), ('85', 'JASHPUR'), ('850', 'Udgir'), ('851', 'Gangakhed'),
             ('852', 'Betul'), ('853', 'Pandhurna'), ('854', 'Multai'), ('855', 'Balaghat'), ('856', 'Chhindwara'),
             ('857', 'Bhimavaram'), ('858', 'Tirupati'), ('859', 'Vijayawada'), ('86', 'SARANGARH'),
             ('860', 'Rajahmundry'), ('861', 'Nalgunda'), ('862', 'KOLKATA'), ('863', 'BANER'),
             ('864', 'MITC- Kandivali'), ('865', 'Nellore'), ('866', 'Rajkot'), ('867', 'Goa'), ('868', 'Ahmedabad'),
             ('869', 'PALGHAR'), ('87', 'KONDAGAON'), ('870', 'Elwala'), ('871', 'Balaganur'), ('872', 'Panchavati'),
             ('873', 'Shirur'), ('878', 'NARSINGHPUR'), ('88', 'NANGAL'), ('888', 'Chickyanachatra'), ('889', 'ELWALA'),
             ('89', 'Udaipur'), ('890', 'CORNAGAL'), ('891', 'Faridabad'), ('892', 'Borivali'), ('893', 'Colaba'),
             ('894', 'Bina'), ('895', 'Bemetara'), ('896', 'GHATKOPAR'), ('897', 'Ramban'), ('898', 'LAJPAT NAGAR'),
             ('899', 'Girgaon CHOWPATTY'), ('90', 'TPT NAGAR (JAIPUR)'), ('900', 'MULUND'), ('901', 'WAZIRPUR'),
             ('902', 'Chafula'), ('903', 'Pimpalgaon'), ('904', 'Bhanupratappur'), ('905', 'BHAGHIRATHUPURA'),
             ('906', 'Tikamgarh'), ('907', 'Rajim'), ('908', 'NARSINGHPUR'), ('909', 'Panna'), ('91', 'PUNCHKULA'),
             ('910', 'Delhi CCC'), ('911', 'Delhi Trans Jamuna'), ('912', 'Jhalawar'), ('913', 'Jaipur'),
             ('914', 'Dausa'), ('915', 'Baran'), ('916', 'SIDDHPUR'), ('917', 'KATHLAL'), ('918', 'LUNAWADA'),
             ('919', 'BORSAD'), ('92', 'DARIYAGANJ'), ('920', 'Ahmedabad'), ('921', 'BAYAD'), ('922', 'HIMATNAGAR'),
             ('923', 'PATAN'), ('924', 'VADGAM'), ('925', 'Danta'), ('926', 'THARAD'), ('927', 'Dakor'),
             ('928', 'Coimbatore'), ('929', 'Mumbai A.O'), ('93', 'SOHNA'), ('930', 'Ghaziabad'), ('931', 'Jalgaon'),
             ('932', 'SITAPUR'), ('933', 'KORBA'), ('934', 'SAVDA'), ('935', 'Baroda'), ('936', 'BHIWADI-2'),
             ('937', 'Ambagarh Chowki'), ('938', 'Sihora'), ('939', 'shakti branch'), ('94', 'Karnal'),
             ('940', 'Baloda Bazar-2'), ('941', 'Jammu'), ('942', 'Chakan (W/o 2nd &4th)'),
             ('943', 'Nashik ( W/o - 2nd & 4th )'), ('944', 'Nagpur 1st &3rd'), ('945', 'Haridwar 2nd&04th'),
             ('946', 'MITC- Kandivali'), ('947', 'Burdwan'), ('948', 'Terakanambi'), ('949', 'Rohru'), ('95', 'DODA'),
             ('950', 'Jaora'), ('951', 'AKHNOOR'), ('952', 'Karimnagar'), ('953', 'Vashi'),
             ('954', 'Kandivali - 1st & 3rd'), ('955', 'Ramnad'), ('956', 'Parbhani'), ('957', 'Vashi'),
             ('958', 'Veraval'), ('959', 'Pandharpur'), ('96', 'SAMBA'), ('960', 'Meerut'), ('961', 'Hubli'),
             ('962', 'Aligarh'), ('963', 'Shahjhanpur'), ('964', 'Basti'), ('965', 'Azamgarah'), ('966', 'Vellor'),
             ('967', 'Naidupeta'), ('968', 'Kurnool'), ('969', 'Tadepalligudem'), ('97', 'Baramulla'),
             ('970', 'Gudiwada'), ('971', 'Himmatnagar'), ('972', 'Bhildi'), ('973', 'Ranasan'), ('974', 'Talod'),
             ('975', 'Bhavnagar'), ('976', 'Deesa'), ('977', 'Sojitra'), ('978', 'Kalol gn'), ('979', 'Bhiloda'),
             ('98', 'BORSAD'), ('980', 'Panthawada'), ('981', 'Khedbrahma'), ('982', 'Bhavnagar'), ('983', 'Ashta'),
             ('984', 'Sonkatch'), ('985', 'Manasa'), ('986', 'Badnawar'), ('987', 'Dhar'), ('988', 'Pachore'),
             ('989', 'Shamgarh'), ('99', 'WAV'), ('990', 'Alote'), ('991', 'Malegaon'), ('992', 'Ambejogai'),
             ('993', 'Bhokar'), ('994', 'Bhoom'), ('995', 'Buldhana'), ('996', 'Degloor'), ('997', 'Ghansawangi'),
             ('998', 'Jafrabad'), ('999', 'Jamner')])
        self.chtext71 = OrderedDict(
            [('01', 'MITC - 05th Flr.'), ('02', 'MITC - 06th Flr.'), ('03', 'AFS'), ('04', 'FES'), ('05', 'MEML'),
             ('05TH FLR.', 'MITC - 05th Floor'), ('06', 'MVML'), ('06TH FLR.', 'MITC-06th Floor'), ('07', 'Chinchwad'),
             ('08', 'SBU'), ('09', 'MRV'), ('10', 'MFCWL'), ('AFS', 'AFS'), ('FES', 'FES'), ('MEML', 'MEML'),
             ('MVML', 'MVML'), ('PUNE-CHINCH', 'Chinchwad'), ('SBU', 'Strategic Business Unit'),
             ('SSU', 'Spare Sourcing Unit'), ('SUB LOC', 'Sub Location')])
        self.chtext72 = OrderedDict([('00', 'Default')])
        self.chtext83 = OrderedDict([('ACC. CO', 'Account Co-ordinator'), ('AI', 'Account Intern'),
                                     ('AMQT', 'Assistant Manager Quality & Training'), ('ASC', 'Associate'),
                                     ('ASC Adm', 'Admin Associate'), ('ASC BUY', 'Associate - Buying'),
                                     ('ASC CC', 'Associate - Customer Care'), ('ASC DP', 'Associate - Demand Planning'),
                                     ('ASC QA', 'Associate - QA Audit'), ('ASC SS', 'Associate - Spare Sourcing'),
                                     ('ASCOP', 'Associate - Operations'), ('ASS Comm', 'Assistant - Commercial'),
                                     ('ASSOCIATE BUSINESS E', 'Associate Business Excellence'),
                                     ('ASST. MAN', 'Asst. Manager'), ('BO', 'Back Office Executive'),
                                     ('C MANAGER', 'Catalouge Manager'), ('CATE. MGR', 'Category Manager'),
                                     ('CBO', 'Chief Business Officer'), ('CCINCRGE', 'Central Commercial In-change'),
                                     ('CEO', 'CEO'), ('CFO', 'Chief Finance Officer'),
                                     ('CITY SUPERVISOR', 'City Supervisor'), ('CONTRACT', 'CONTRACT'),
                                     ('CRE', 'Customer Relation Executive'), ('CSA', 'Customer Service Associate'),
                                     ('CUST CE', 'Customer - Care Executive'), ('DA', 'Data Analyst'),
                                     ('DEMONSTRATOR', 'Demonstrator'), ('DEPUTY MA', 'Deputy Manager'),
                                     ('EX- COM', 'Executive - Commercial'), ('EXE', 'Executive'),
                                     ('EXE.Q&M', 'Executive-Quality & Maintenance'),
                                     ('EXEC-ADMIN', 'Executive - Administration'), ('EXEC-IT', 'Executive - IT'),
                                     ('EXECU-ACC', 'Executive - Accounts'),
                                     ('Exe OL', 'Executive - Outbound Logistics'),
                                     ('Exe PS', 'Executive - Project and System'), ('Exe QA', 'Executive - Quality'),
                                     ('Exe SBU', 'Executive - SBU'), ('Exe WO', 'Executive - Warehouse Operations'),
                                     ('FE', 'Field Executive'), ('GM / VP', 'GM / VP'),
                                     ('GMB / SEC', 'GMB / Sector President'), ('GRAP-DESIGN', 'Graphic Designer'),
                                     ('HEAD', 'HEAD-MBPO'), ('HEAD F & A', 'Head F&A Practice'),
                                     ('HEAD OF D', 'Head of Dept'), ('HEAD- OP. ', 'Head - Operations'),
                                     ('HR HEAD', 'Head HR & HR Practices'), ('JR. EXECU', 'Jr. Executive'),
                                     ('JR.EXE HR', 'Junior Executive- HR'), ('LAC', 'Lead Account Co-ordinator'),
                                     ('LAUNCH MANAGER', 'Launch Manager'), ('LEAD CE', 'Lead Contact Centre'),
                                     ('MANAGER', 'Manager'), ('MH SR', 'Mahindra Sales Represantative'),
                                     ('MIS Exe', 'MIS executive'), ('MKT SR', 'Market Sales Representative'),
                                     ('MSR', 'MSR'), ('Med Att', 'Medical Attendant'), ('Med Off', 'Medical Officer'),
                                     ('OA', 'Office Assistant'), ('OP. HEAD', 'Operational Head'),
                                     ('OP. SUP', 'Operational Supervisor'), ('OP.EXEC', 'Operational Executive'),
                                     ('PE', 'Purchase Engineer'), ('PM', 'Process Manager'),
                                     ('POO', 'Photo Operations Officer'),
                                     ('PROCUREMENT SUPPORT', 'PROCUREMENT SUPPORT'),
                                     ('PROGRAM MGR.', 'PROGRAM MANAGER'), ('QA Off', 'QA Officer'),
                                     ('QA-COACH', 'Quality Coach'), ('QA-I', 'QA Intern'),
                                     ('QUALITY-ANA', 'Quality Analyst'), ('QUERY-CO', 'Query Co-ordinator'),
                                     ('Qua Aud', 'Quality Auditor'), ('Qua Eng', 'Quality Engineer'),
                                     ('Qua Insp', 'Quality Inspector'), ('REGIONAL MANGR', 'Regional Manager'),
                                     ('SAL. EXECU', 'Sales Executive'), ('SALES CONS', 'Sales Consultant'),
                                     ('SCRE', 'Senior Customer Relationship Executive'), ('SE', 'Supply Executive'),
                                     ('SER.TECH', 'Service Technician - Tractor'),
                                     ('SERVICE TECH', 'Service Technician - Implements'),
                                     ('SERVICE TECHNICIAN', 'Service Technician'),
                                     ('SERVICE TECHNICIAN -', 'Service Technician - Implements'),
                                     ('SH.SQ', 'Shift Supervisor- Quality'), ('SME', 'Subject Matter Expert'),
                                     ('SP', 'Seconadry packing'), ('SR', 'Sales Representative'),
                                     ('SR EX ACC', 'Sr. Executive - Accounts'),
                                     ('SR EX ADMIN', 'Sr. Executive - Administration'),
                                     ('SR EX MIS', 'Sr. Executive - MIS'), ('SR. EXECU', 'Sr. Executive'),
                                     ('SR. MANAG', 'Sr. Manager'), ('SR.ASC', 'Sr.Associate'),
                                     ('SR.WRITER', 'Senior Writer'), ('TC', 'Team Coach'),
                                     ('TEAM-QUALITY', 'Team Leader - QUALITY'), ('TL', 'Team  Leader'),
                                     ('TR.EXEC', 'Training Executive'), ('TRAINEE', 'TRAINEE'), ('TRAINER', 'Trainer'),
                                     ('Tech Train Mgr', 'Technical Training Manager'),
                                     ('Trn Co', 'Training Coordinator'), ('UNDEFINED', 'UNDEFINED'),
                                     ('WE', 'Warehouse Executive'), ('Zon Co', 'Zonal Coordinator')])
        self.chtext84 = OrderedDict(
            [('ACCPY', 'Accounts Payable'), ('ACCPY - GST', 'Accounts Payable - GST'), ('ADMIN', 'Administration'),
             ('BD', 'Business Development'), ('BFSI', 'BFSI'), ('CBO', 'CBO - OFFICE'), ('CENG', 'Customer Engagement'),
             ('CEO', 'CEO Office'), ('COMP', 'Compliance'), ('D&B', 'D&B'), ('DIG', 'Digitization'),
             ('EMPBENEFIT', 'Employee Benefits'), ('ER&D', 'Employee Relation and Development'),
             ('FDSTAFF', 'Farm Division- Staffing'), ('FIN', 'Finance'), ('HR', 'Human Resource'),
             ('HR-ER', 'HR ER Client Deputed '), ('IT', 'Information Technology'), ('KYC-P', 'KYC-Process'),
             ('MPWROL', 'MAHINDRA POWEROL'), ('NI', 'New Initiatives'), ('OPS', 'Opeartions'), ('POWEROL', 'POWEROL'),
             ('RETAIL', 'Retail'), ('SALP', 'Salary Processing'), ('STAFF', 'STAFFING'), ('UNDEF', 'UNDEFINED')])
        self.chtext85 = OrderedDict(
            [('00', 'General'), ('01', 'Recruitment'), ('02', 'Onboarding'), ('03', 'Staffing'), ('04', 'Operations'),
             ('05', 'Kiran team'), ('06', 'Sales'), ('07', 'Alturas'), ('08', 'AS-Marketing'), ('09', 'M2ALL'),
             ('10', 'Jawa'), ('11', 'XUV300'), ('12', 'Marazzo'), ('13', 'Demonstrator - Dhruv'),
             ('14', 'Demonstrator - Tractor'), ('15', 'Quality'), ('16', 'MIS'), ('17', 'Manthan'), ('18', 'Digital'),
             ('19', 'Sparsh'), ('20', 'MRHFL'), ('21', 'MFCSL-SMR'), ('22', 'SCM-Backend'), ('23', 'MFCSL-PSF'),
             ('24', 'MLDL'), ('25', 'SCM-Sales')])
        self.chtext88 = OrderedDict([('01', 'M&M Accounts HO'), ('02', 'HOUZZ'), ('03', 'MRHFL'), ('04', 'Smartshift'),
                                     ('05', 'FD-Sales, Customer and Channel Care'), ('06', 'MIBL'), ('07', 'M2ALL'),
                                     ('08', 'M&M IT'), ('09', 'MAHINDRA WORLD UNIVERSITY'), ('10', 'D&B'),
                                     ('100', 'Mahindra & Mahindra - FD (CME)'),
                                     ('101', 'Mahindra & Mahindra- Auto Sector(Igatpuri)'),
                                     ('102', 'Mahindra & Mahindra- Auto Sector(Chakan)'),
                                     ('103', 'Mahindra & Mahindra - AD (Sales,MKTG & Cust)'),
                                     ('11', 'AS Marketing- M2all '), ('12', 'AS Marketing-Accounts'),
                                     ('13', 'Flipkart'), ('14', 'M & M IT -SHIVA'), ('15', 'M&M AUTO passon'),
                                     ('16', 'M&M FD'), ('17', 'Mahindra & Mahindra Financial Services Ltd'),
                                     ('18', 'Mahindra Agri (Grapes Seasonal Business)   (New)'),
                                     ('19', 'Mahindra Home Finance '), ('20', 'MIBS'), ('21', 'MRHFL- scanning'),
                                     ('22', 'MTBD'), ('23', 'MTWL'), ('24', 'Powerol'), ('25', 'SBU'),
                                     ('26', 'Smartshift- Accounts'), ('27', 'SSU'), ('28', 'SABARO'),
                                     ('29', 'M & M IT-SAURABH YASHWANT MULIK'), ('30', 'MVML'),
                                     ('31', 'Mahindra Agri Business Solutions'), ('32', 'Trringo'),
                                     ('33', 'Farm Division'), ('34', 'Auto Division'), ('35', 'Truck and Bus Division'),
                                     ('36', 'Msolve IT'), ('37', 'MRHFL(S5S)'),
                                     ('38', 'M&M Auto division ( Launch Manager)'), ('39', 'MLDL'),
                                     ('40', 'Auto Division ( L Vasudevan)'), ('41', 'M&M Admin'), ('42', 'MLDL-2'),
                                     ('43', 'M&M HO Accounts- Gupta Harish'),
                                     ('44', ' M&M - AD - Sales - Preowned Car Bus. - Richa'), ('45', 'M&M- Neelkamal'),
                                     ('46', 'MFCSL'), ('47', 'MMFSL'), ('48', 'FD-Supply chain-central prodction'),
                                     ('49', 'Mahindra Finance- KYC Process'), ('50', 'IIC - Mahindra & Mahindra'),
                                     ('51', 'ER&D DEPT - KND M&M'), ('52', 'M&M Construction Equipment'),
                                     ('53', 'MLDL- AP'), ('54', 'M&M Two Wheeler Div'), ('55', 'MTWD'),
                                     ('56', 'Myntra'), ('57', 'Powerol (BSNL)'), ('58', 'M&M AFS'),
                                     ('59', 'Mahindra & Mahindra Common Services'),
                                     ('60', 'M&M Auto Division (Cyril Patole)'), ('61', 'M&M Neelkamal'),
                                     ('62', 'Mahindra and Mahindra Aerospace'), ('63', 'MVARTA'),
                                     ('64', 'M&M T.D Canteen'), ('65', 'M&M Marketing'), ('66', 'Mahindra Samrudhhi'),
                                     ('67', 'M&M ER&D'), ('68', 'MLDL-IT-Trainee'), ('69', 'MVML-IT'), ('70', 'MAPL'),
                                     ('71', 'Gipps Aero'), ('72', 'Mahindra Electric'),
                                     ('73', 'Mahindra & Mahindra -AD'), ('74', 'Classic Legends Pvt Ltd.-MTBD'),
                                     ('75', 'MASL'), ('76', 'Mahindra Agri Business Solutions - Grapes'),
                                     ('77', 'M&M -FD Div.Service Tech.'), ('78', 'M&M Tool and Die Plant '),
                                     ('79', 'MIBL Worli'), ('80', 'SBU-IT'), ('81', 'MLDL-Saurabh'),
                                     ('82', 'NBS International ltd '), ('83', 'MGPL'),
                                     ('84', 'Mahindra & Mahindra Farm Equipment Sector'),
                                     ('85', 'M&M-FES Farm Machinery'), ('86', 'Chennai-MRV (IT)'),
                                     ('87', 'M & M - International Operation'),
                                     ('88', 'Mahindra & Mahindra Auto Sector'),
                                     ('89', 'M & M GCO- Compensation and Benefits'),
                                     ('90', 'Mahindra-Precision Farming'),
                                     ('91', 'Mahindra & Mahindra - Haridwar (Admin)'),
                                     ('92', 'M&M-FES Farm Machinery (TCD)'), ('93', 'M&M - AD Kandivali Accounts'),
                                     ('94', 'Bharucha & Partners'), ('95', 'M & M AFS (Worli)'), ('96', 'MIBL Solapur'),
                                     ('97', 'Mahindra & Mahindra Auto Farm Sector'),
                                     ('98', 'Mahindra Automotive Sector'),
                                     ('99', 'Mahindra & Mahindra - Commercial Executive'),
                                     ('129', 'MIBS Staffing')])
        self.chtext89 = OrderedDict([("AP", "ANDHRA PRADESH"), ("AS", "ASSAM"), ("BH", "BIHAR"), ("CG", "CHHATISGARH"), ("CH", "CHANDIGARH"), ("DD", "DAMAN & DIU"), ("DL", "DELHI"), ("DN", "DADRA & NAGAR HAVELI"), ("GA", "GOA"), ("GJ", "GUJARAT"), ("HP", "HIMACHAL PRADESH"), ("HR", "HARYANA"), ("JK", "JAMMU & KASHMIR"), ("JZ", "JHARKHAND"), ("KL", "KERALA"), ("KN", "KARNATAKA"), ("LD", "LAKSHWADEEP"), ("MG", "MEGHALAYA"), ("MH", "MAHARASHTRA"), ("MN", "MANIPUR"), ("MP", "MADHYA PRADESH"), ("MZ", "MIZORAM"), ("NL", "NAGALAND"), ("OR", "ODISHA"), ("PB", "PUNJAB"), ("PN", "PONDICHERRY"), ("RJ", "RAJASTHAN"), ("SK", "SIKKIM"), ("TN", "TAMIL NADU"), ("TR", "TRIPURA"), ("TS", "TELANGANA"), ("UK", "UTTARAKHAND"), ("UP", "UTTAR PRADESH"), ("WB", "WEST BENGAL")])
        self.choice1x = ["Please Select"]
        self.choice1x.extend(list(self.chtext1.keys()))
        self.choice63x = ["Please Select"]
        self.choice63x.extend([F"{a} ({b})" for a,b in zip(self.chtext63.keys(),self.chtext63.values())])
        self.choice64x = ["Please Select"]
        self.choice64x.extend([F"{a} ({b})" for a,b in zip(self.chtext64.keys(),self.chtext64.values())])
        self.choice65x = ["Please Select"]
        self.choice65x.extend([F"{a} ({b})" for a,b in zip(self.chtext65.keys(),self.chtext65.values())])
        self.choice66x = ["Please Select"]
        self.choice66x.extend([F"{a} ({b})" for a,b in zip(self.chtext66.keys(),self.chtext66.values())])
        self.choice68x = ["Please Select"]
        self.choice68x.extend([F"{a} ({b})" for a,b in zip(self.chtext68.keys(),self.chtext68.values())])
        self.choice69x = ["Please Select"]
        self.choice69x.extend([F"{a} ({b})" for a,b in zip(self.chtext69.keys(),self.chtext69.values())])
        self.choice70x = ["Please Select"]
        self.choice70x.extend([F"{a} ({b})" for a,b in zip(self.chtext70.keys(),self.chtext70.values())])
        self.choice71x = ["Please Select"]
        self.choice71x.extend([F"{a} ({b})" for a,b in zip(self.chtext71.keys(),self.chtext71.values())])
        self.choice72x = ["Please Select"]
        self.choice72x.extend([F"{a} ({b})" for a,b in zip(self.chtext72.keys(),self.chtext72.values())])
        self.choice83x = ["Please Select"]
        self.choice83x.extend([F"{a} ({b})" for a,b in zip(self.chtext83.keys(),self.chtext83.values())])
        self.choice84x = ["Please Select"]
        self.choice84x.extend([F"{a} ({b})" for a,b in zip(self.chtext84.keys(),self.chtext84.values())])
        self.choice85x = ["Please Select"]
        self.choice85x.extend([F"{a} ({b})" for a,b in zip(self.chtext85.keys(),self.chtext85.values())])
        self.choice88x = ["Please Select"]
        self.choice88x.extend([F"{a} ({b})" for a,b in zip(self.chtext88.keys(),self.chtext88.values())])
        self.choice89x = ["Please Select"]
        self.choice89x.extend([F"{a} ({b})" for a,b in zip(self.chtext89.keys(),self.chtext89.values())])

        self.choice1 = ["Please Select"]
        self.choice1.extend(self.chtext1.keys())
        self.choice63 = ["Please Select"]
        self.choice63.extend(self.chtext63.keys())
        self.choice64 = ["Please Select"]
        self.choice64.extend(self.chtext64.keys())
        self.choice65 = ["Please Select"]
        self.choice65.extend(self.chtext65.keys())
        self.choice66 = ["Please Select"]
        self.choice66.extend(self.chtext66.keys())
        self.choice68 = ["Please Select"]
        self.choice68.extend(self.chtext68.keys())
        self.choice69 = ["Please Select"]
        self.choice69.extend(self.chtext69.keys())
        self.choice70 = ["Please Select"]
        self.choice70.extend(self.chtext70.keys())
        self.choice71 = ["Please Select"]
        self.choice71.extend(self.chtext71.keys())
        self.choice72 = ["Please Select"]
        self.choice72.extend(self.chtext72.keys())
        self.choice83 = ["Please Select"]
        self.choice83.extend(self.chtext83.keys())
        self.choice84 = ["Please Select"]
        self.choice84.extend(self.chtext84.keys())
        self.choice85 = ["Please Select"]
        self.choice85.extend(self.chtext85.keys())
        self.choice88 = ["Please Select"]
        self.choice88.extend(self.chtext88.keys())
        self.choice89 = ["Please Select"]
        self.choice89.extend(self.chtext89.keys())

        self.text1 = wx.Choice(self.p, -1, choices=list(self.choice1x))
        self.text2 = wx.TextCtrl(self.p)
        self.text3 = wx.TextCtrl(self.p)
        self.text4 = wx.TextCtrl(self.p)
        self.text5 = wx.TextCtrl(self.p)
        self.text6 = wx.TextCtrl(self.p)
        self.text7 = wx.TextCtrl(self.p)
        self.text8 = wx.TextCtrl(self.p)
        self.text9 = wx.TextCtrl(self.p)
        self.text10 = wx.TextCtrl(self.p)
        self.text11 = wx.TextCtrl(self.p)
        self.text12 = wx.TextCtrl(self.p)
        self.text13 = wx.TextCtrl(self.p)
        self.text14 = wx.TextCtrl(self.p)
        self.text15 = wx.TextCtrl(self.p)
        self.text16 = wx.TextCtrl(self.p)
        self.text17 = wx.TextCtrl(self.p)
        self.text18 = wx.TextCtrl(self.p)
        self.text19 = wx.TextCtrl(self.p)
        self.text20 = wx.TextCtrl(self.p)
        self.text21 = wx.TextCtrl(self.p)
        self.text22 = wx.TextCtrl(self.p)
        self.text23 = wx.TextCtrl(self.p)
        self.text24 = wx.TextCtrl(self.p)
        self.text25 = wx.TextCtrl(self.p)
        self.text26 = wx.TextCtrl(self.p)
        self.text27 = wx.TextCtrl(self.p)
        self.text28 = wx.TextCtrl(self.p)
        self.text29 = wx.TextCtrl(self.p)
        self.text30 = wx.TextCtrl(self.p)
        self.text31 = wx.TextCtrl(self.p)
        self.text32 = wx.TextCtrl(self.p)
        self.text33 = wx.TextCtrl(self.p)
        self.text34 = wx.TextCtrl(self.p)
        self.text35 = wx.TextCtrl(self.p)
        self.text36 = wx.TextCtrl(self.p)
        self.text37 = wx.TextCtrl(self.p)
        self.text38 = wx.TextCtrl(self.p)
        self.text39 = wx.TextCtrl(self.p)
        self.text40 = wx.TextCtrl(self.p)
        self.text41 = wx.TextCtrl(self.p)
        self.text42 = wx.TextCtrl(self.p)
        self.text43 = wx.TextCtrl(self.p)
        self.text44 = wx.TextCtrl(self.p)
        self.text45 = wx.TextCtrl(self.p)
        self.text46 = wx.TextCtrl(self.p)
        self.text47 = wx.TextCtrl(self.p)
        self.text48 = wx.TextCtrl(self.p)
        self.text49 = wx.TextCtrl(self.p)
        self.text50 = wx.TextCtrl(self.p)
        self.text51 = wx.TextCtrl(self.p)
        self.text52 = wx.TextCtrl(self.p)
        self.text53 = wx.TextCtrl(self.p)
        self.text54 = wx.TextCtrl(self.p)
        self.text55 = wx.TextCtrl(self.p)
        self.text56 = wx.TextCtrl(self.p)
        self.text57 = wx.TextCtrl(self.p)
        self.text58 = wx.TextCtrl(self.p)
        self.text59 = wx.TextCtrl(self.p)
        self.text60 = wx.TextCtrl(self.p)
        self.text61 = wx.TextCtrl(self.p)
        self.text62 = wx.TextCtrl(self.p)
        self.text63 = wx.Choice(self.p, -1, choices=self.choice63x)
        self.text64 = wx.Choice(self.p, -1, choices=self.choice64x)
        self.text65 = wx.Choice(self.p, -1, choices=self.choice65x)
        self.text66 = wx.Choice(self.p, -1, choices=self.choice66x)
        self.text67 = wx.TextCtrl(self.p)
        self.text68 = wx.Choice(self.p, -1, choices=self.choice68x)
        self.text69 = wx.Choice(self.p, -1, choices=self.choice69x)
        self.text70 = wx.Choice(self.p, -1, choices=self.choice70x)
        self.text71 = wx.Choice(self.p, -1, choices=self.choice71x)
        self.text72 = wx.Choice(self.p, -1, choices=self.choice72x)
        self.text73 = wx.TextCtrl(self.p)
        self.text74 = wx.TextCtrl(self.p)
        self.text75 = wx.TextCtrl(self.p)
        self.text76 = wx.TextCtrl(self.p)
        self.text77 = wx.TextCtrl(self.p)
        self.text78 = wx.TextCtrl(self.p)
        self.text79 = wx.TextCtrl(self.p)
        self.text80 = wx.TextCtrl(self.p)
        self.text81 = wx.TextCtrl(self.p)
        self.text82 = wx.TextCtrl(self.p)
        self.text83 = wx.Choice(self.p, -1, choices=self.choice83x)
        self.text84 = wx.Choice(self.p, -1, choices=self.choice84x)
        self.text85 = wx.Choice(self.p, -1, choices=self.choice85x)
        self.text86 = wx.TextCtrl(self.p)
        self.text87 = wx.TextCtrl(self.p)
        self.text88 = wx.Choice(self.p, -1, choices=self.choice88x)
        self.text89 = wx.Choice(self.p, -1, choices=self.choice89x)
        self.text90 = wx.TextCtrl(self.p)
        self.text91 = wx.TextCtrl(self.p)

        font = wx.SystemSettings.GetFont(wx.SYS_SYSTEM_FONT)
        font.SetPointSize(80)

        self.sendmail_button = wx.Button(self.p, 3, label="SEND MAILS", size=(140, 60))
        self.sendmail_button.SetFont(font)
        self.sendmail_button.SetBackgroundColour("#1395FA")

        self.save_button = wx.Button(self.p, 1, label="SAVE", size=(140, 60))
        self.save_button.SetFont(font)
        self.save_button.SetBackgroundColour("#1395FA")

        # self.submit_button = wx.Button(self.p, 2, label="SUBMIT", size=(240, 50))
        # self.submit_button.SetFont(font)
        # self.submit_button.SetBackgroundColour("#1395FA")
        self.sendmail_button.Disable()

        # self.sendmail_button.SetBackgroundColour("#0B0B3B")
        # self.Bind(wx.EVT_BUTTON, self.submit_all, self.submit_button)
        self.Bind(wx.EVT_BUTTON, self.get_save_data, self.save_button)
        self.Bind(wx.EVT_BUTTON, self.mail_send, self.sendmail_button)

        font = wx.Font(14, wx.ROMAN, wx.NORMAL, wx.LIGHT)

        mappings = {1: "Salutation/Title", 2: "First Name", 3: "Middle Name", 4: "Last Name", 5: "Gender",
                    6: "Father’s/Husband Name", 7: "Employee Relation", 8: "Employee Name", 9: "Marital Status",
                    10: "Spouse Name", 11: "No. of Children", 12: "Present Address", 13: "Present City",
                    14: "Present State", 15: "Present Pin code", 16: "Present Phone", 17: "Permanent Address",
                    18: "Permanent City", 19: "Permanent State", 20: "Permanent Pin code", 21: "Permanent Phone",
                    22: "Primary Bank Name", 23: "Primary IFSC Code", 24: "Primary Account No",
                    25: "Primary name as per bank", 26: "MICR", 27: "Expected Date of Joining", 28: "Date of Birth",
                    29: "Qualification", 30: "Blood Group", 31: "Emergency Phone No", 32: "Emergency Contact Person",
                    33: "Aadhar card no", 34: "Permanent Account No", 35: "Email Address", 36: "Personal Mobile No",
                    37: "Establishment Address", 38: "Universal Account Number", 39: "PF Account Number",
                    40: "Date of joining (Unexempted)", 41: "Date of exit (Unexempted)",
                    42: "Scheme Certificate (if issued)", 43: "PPO Number (if issued)",
                    44: "Non-Contributory Period (NCP) Days", 45: "Name & Address of the Trust", 46: "UAN",
                    47: "Member EPS A/c No", 48: "Date of joining (Exempted)", 49: "Date of exit (Exempted)",
                    50: "Scheme Certificate No (if issued)", 51: "Non Contributory period (NCP) Days",
                    52: "State Country of origin", 53: "Passport No", 54: "Validity of Passport", 55: "Upload Passport Photo",
                    56: "Upload Docs", 57: "Employee Number", 58: "Confirm Date of Joining", 59: "Training From",
                    60: "Date of Probation", 61: "Date of Retirement", 62: "Date of confirmation", 63: "Payroll Code",
                    64: "Category Code", 65: "Status Code", 66: "Grade Code", 67: "Designation", 68: "Cost Center Code",
                    69: "Business Area Code", 70: "Location Code",71: "Sub-Location Code",72: "Occupation Code",73: "PF Registration Code",74: "PF wef Dt",75: "E.S.I. No.",76: "Reports to emp",77: "Reporting Manager's Token ID",78: "Mobile number",79: "Web user name",80: "Web user password",81: "Web access level",82: "Attendance user id",83: "Designation Code",84: "Department Code",85: "Sub-Department Code",86: "CR Mem No",87: "Client name",88: "Client name code",89: "State code",90: "Group joining date",91: "Reporting Manager Email id"}
        dictmap = {62:self.chtext63,63:self.chtext64,64:self.chtext65,65:self.chtext66,67:self.chtext68,68:self.chtext69,69:self.chtext70,70:self.chtext71,71:self.chtext72,82:self.chtext83,83:self.chtext84,84:self.chtext85,87:self.chtext88,88:self.chtext89}
        # for i, x in enumerate([self.text1,self.text2,self.text3,self.text4,self.text5,self.text6,self.text7,self.text8,self.text9,self.text10,self.text11,self.text12,self.text13,self.text14,self.text15,self.text16,self.text17,self.text18,self.text19,self.text20,self.text21,self.text22,self.text23,self.text24,self.text25,self.text26,self.text27,self.text28,self.text29,self.text30,self.text31,self.text32,self.text33,self.text34,self.text35,self.text36,self.text37,self.text38,self.text39,self.text40,self.text41,self.text42,self.text43,self.text44,self.text45,self.text46,self.text47,self.text48,self.text49,self.text50,self.text51,self.text52,self.text53,self.text54,self.text55,self.text56,self.text57,self.text58,self.text59,self.text60,self.text61,self.text62,self.text63,self.text64,self.text65,self.text66,self.text67,self.text68,self.text69,self.text70,self.text71,self.text72,self.text73,self.text74,self.text75,self.text76,self.text77,self.text78,self.text79,self.text80,self.text81,self.text82,self.text83,self.text84,self.text85,self.text86,self.text87,self.text88]):
        for i, x in enumerate([self.text1,self.text2,self.text3,self.text4,self.text5,self.text6,self.text7,self.text8,self.text9,self.text10,self.text11,self.text12,self.text13,self.text14,self.text15,self.text16,self.text17,self.text18,self.text19,self.text20,self.text21,self.text22,self.text23,self.text24,self.text25,self.text26,self.text27,self.text28,self.text29,self.text30,self.text31,self.text32,self.text33,self.text34,self.text35,self.text36,self.text37,self.text38,self.text39,self.text40,self.text41,self.text42,self.text43,self.text44,self.text45,self.text46,self.text47,self.text48,self.text49,self.text50,self.text51,self.text52,self.text53,self.text54,self.text55,self.text56,self.text57,self.text58,self.text59,self.text60,self.text61,self.text62,self.text63,self.text64,self.text65,self.text66,self.text67,self.text68,self.text69,self.text70,self.text71,self.text72,self.text73,self.text74,self.text75,self.text76,self.text77,self.text78,self.text79,self.text80,self.text81,self.text82,self.text83,self.text84,self.text85,self.text86,self.text87,self.text88,self.text89,self.text90,self.text91]):
            if self.data.get(mappings[i+1]) is not None:
                if x in [self.text1,self.text63,self.text64,self.text65,self.text66,self.text68,self.text69,self.text70,self.text71,self.text72,self.text83,self.text84,self.text85,self.text88,self.text89]:
                    if self.data.get(mappings[i+1]) and i != 0:
                        x.SetSelection(list(dictmap[i].values()).index(str(self.data.get(mappings[i+1])).split("[")[0])+1)
                    elif i == 0:
                        x.SetSelection(list(self.choice1x).index(self.data.get(mappings[i+1])))
                    else:
                        x.SetSelection(0)
                elif i == 72:
                    if not self.data.get(mappings[i+1]):
                        x.SetValue("KDMAL0213914000")
                        # x.SetValue("KDMAL0213914")
                    else:
                        x.SetValue(str(self.data.get(mappings[i + 1])))
                elif i == 80:
                    if not self.data.get(mappings[i+1]):
                        x.SetValue("00")
                    else:
                        x.SetValue(str(self.data.get(mappings[i + 1])))
                else:
                    x.SetValue(str(self.data.get(mappings[i+1])))

        r = 12
        if self.data["Permanent-Address"] == "Same As Above":
            for x in [self.text17,self.text18,self.text19,self.text20,self.text21]:
                x.SetValue(str(self.data.get(mappings[r])))
                r += 1
        if self.data["Marital Status"] == "Unmarried":
            self.text10.SetValue("")
            self.text11.SetValue("")
            # self.text10.SetBackgroundColour("#222223")
            # self.text10.Disable()
            self.text11.Disable()

        if self.data['Choose One'] == "Exempted":
            self.text37.SetValue("")
            self.text37.Disable()
            self.text38.SetValue("")
            self.text38.Disable()
            self.text39.SetValue("")
            self.text39.Disable()
            self.text40.SetValue("")
            self.text40.Disable()
            self.text41.SetValue("")
            self.text41.Disable()
            self.text42.SetValue("")
            self.text42.Disable()
            self.text43.SetValue("")
            self.text43.Disable()
            self.text44.SetValue("")
            self.text44.Disable()
        elif self.data['Choose One'] == "Un-exempted":
            self.text45.SetValue("")
            self.text45.Disable()
            self.text46.SetValue("")
            self.text46.Disable()
            self.text47.SetValue("")
            self.text47.Disable()
            self.text48.SetValue("")
            self.text48.Disable()
            self.text49.SetValue("")
            self.text49.Disable()
            self.text50.SetValue("")
            self.text50.Disable()
            self.text51.SetValue("")
            self.text51.Disable()

        if self.data["Worked"] == "No":
            self.text52.SetValue("")
            self.text52.Disable()
            self.text53.SetValue("")
            self.text53.Disable()
            self.text54.SetValue("")
            self.text54.Disable()


        for i, x in enumerate([self.label1,self.text1,self.label2,self.text2,self.label3,self.text3,self.label4,self.text4,self.label5,self.text5,self.label6,self.text6,self.label7,self.text7,self.label8,self.text8,self.label9,self.text9,self.label10,self.text10,self.label11,self.text11,self.label12,self.text12,self.label13,self.text13,self.label14,self.text14,self.label15,self.text15,self.label16,self.text16,self.label17,self.text17,self.label18,self.text18,self.label19,self.text19,self.label20,self.text20,self.label21,self.text21,self.label22,self.text22,self.label23,self.text23,self.label24,self.text24,self.label25,self.text25,self.label26,self.text26,self.label27,self.text27,self.label28,self.text28,self.label29,self.text29,self.label30,self.text30,self.label31,self.text31,self.label32,self.text32,self.label33,self.text33,self.label34,self.text34,self.label35,self.text35,self.label36,self.text36,self.label37,self.text37,self.label38,self.text38,self.label39,self.text39,self.label40,self.text40,self.label41,self.text41,self.label42,self.text42,self.label43,self.text43,self.label44,self.text44,self.label45,self.text45,self.label46,self.text46,self.label47,self.text47,self.label48,self.text48,self.label49,self.text49,self.label50,self.text50,self.label51,self.text51,self.label52,self.text52,self.label53,self.text53,self.label54,self.text54,self.label55,self.text55,self.label56,self.text56,self.label88,self.text88,self.label57,self.text57,self.label58,self.text58,self.label59,self.text59,self.label60,self.text60,self.label61,self.text61,self.label62,self.text62,self.label63,self.text63,self.label64,self.text64,self.label65,self.text65,self.label66,self.text66,self.label67,self.text67,self.label68,self.text68,self.label69,self.text69,self.label70,self.text70,self.label71,self.text71,self.label72,self.text72,self.label73,self.text73,self.label74,self.text74,self.label75,self.text75,self.label76,self.text76,self.label77,self.text77,self.label78,self.text78,self.label79,self.text79,self.label80,self.text80,self.label81,self.text81,self.label82,self.text82,self.label83,self.text83,self.label84,self.text84,self.label85,self.text85,self.label86,self.text86,self.label87,self.text87,self.label89,self.text89,self.label90,self.text90,self.label91,self.text91]):
            if i % 2 == 0:
                x.SetFont(font)
            gs.Add(x, 0, wx.EXPAND)

        # gs.Add(None, 0, wx.EXPAND)
        # gs.Add(self.submit_button, 0, wx.EXPAND)

        # mainSizer.Add(gs, 0, wx.EXPAND | wx.GROW)
        # self.p.SetSizer(mainSizer)
        #############################################################3
        self.l1 = wx.StaticText(self.p, label="                                          ")
        self.l2 = wx.StaticText(self.p, label="                                  ")
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(self.l1,0,wx.CENTER | wx.ALL, 40)
        btn_sizer.Add(self.save_button, 0, wx.CENTER | wx.ALL, 40)
        btn_sizer.Add(self.l2,0,wx.CENTER | wx.ALL, 40)
        btn_sizer.Add(self.sendmail_button, 0, wx.CENTER | wx.ALL, 40)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(gs, 1, wx.EXPAND)
        sizer.Add(btn_sizer)
        self.p.SetSizer(sizer)
        self.text79.Bind(wx.EVT_TEXT, self.OnText79)
        self.text57.Bind(wx.EVT_TEXT, self.OnText57)
        self.text58.Bind(wx.EVT_TEXT, self.OnText58)
        self.text90.Bind(wx.EVT_TEXT, self.OnText90)
        self.text88.Bind(wx.EVT_CHOICE, self.OnTextclientcode)


    def OnTextclientcode(self, event):
        self.temp = copy.deepcopy(self.codedict)
        if self.codeseries.get(self.choice88[self.text88.GetSelection()]):
            self.temp[self.codeseries.get(self.choice88[self.text88.GetSelection()])] += 1
            self.text57.ChangeValue(self.codeseries.get(self.choice88[self.text88.GetSelection()])+str(self.temp.get(self.codeseries.get(self.choice88[self.text88.GetSelection()]))))
            self.text79.ChangeValue(self.codeseries.get(self.choice88[self.text88.GetSelection()])+str(self.temp.get(self.codeseries.get(self.choice88[self.text88.GetSelection()]))))
        else:
            self.text57.ChangeValue("-")
            self.text57.ChangeValue("-")

    def OnText79(self, event):
        self.text57.ChangeValue(event.GetString())
    def OnText57(self, event):
        self.text79.ChangeValue(event.GetString())
    def OnText58(self, event):
        self.text90.ChangeValue(event.GetString())
    def OnText90(self, event):
        self.text58.ChangeValue(event.GetString())

    def check_records(self,record):

        flag = "unchecked"
        if record[23] and (len(record[23]) != 11 or not record[23][:4].isalpha() or not record[23][4:].isdigit()):
            flag = "checked"
            pyautogui.alert("Please Check your IFSC Code.")
        elif record[26] and len(record[26]) != 9:
            flag = "checked"
            pyautogui.alert("Please Check your MICR.")

        elif record[27]:
            dn = datetime.datetime.now()
            dn=datetime.datetime.strptime(F"{dn.month}/{dn.day}/{dn.year}", "%m/%d/%Y")
            dje = datetime.datetime.strptime(record[27], "%m/%d/%Y")
            if dn>dje:
                flag = "checked"
                pyautogui.alert("Expected Date of Joining should be greater than today's Date")
        elif record[58]:
            dn = datetime.datetime.now()
            dn= datetime.datetime.strptime(F"{dn.month}/{dn.day}/{dn.year}", "%m/%d/%Y")
            djc = datetime.datetime.strptime(record[58], "%m/%d/%Y")
            if dn>djc:
                flag = "checked"
                pyautogui.alert("Confirm Date of Joining should be greater than today's Date")


        elif record[33] and (len(record[33]) != 12 or not record[33].isdigit()):
            flag = "checked"
            pyautogui.alert("Please Check your Aadhar Card No.")
        elif record[34] and (len(record[34]) != 10 or record[34][3].lower() != "p" or not (record[34][:5]+record[34][-1:]).isalpha() or not record[34][5:9].isdigit()):
            flag = "checked"
            pyautogui.alert("Please Check your Permanent Account Number")
        elif record[46] and (len(record[46]) != 12 or not record[46].isdigit()):
            flag = "checked"
            pyautogui.alert("Please Check your UAN")
        elif record[78] and (len(record[78]) != 10 or not record[78].isdigit()):
            flag = "checked"
            pyautogui.alert("Please Check your Mobile number")
        elif record[28]:
            db = datetime.datetime.strptime(record[28],"%m/%d/%Y")  #09/04/1995
            dn = datetime.datetime.now()                            #05/10/2019
            if dn.year - db.year <= 18:
                if dn.year - db.year == 18:
                    if db.month - dn.month > 0:
                        flag = "checked"
                        pyautogui.alert("Age is Less than 18 Years.")
                    elif db.month - dn.month == 0:
                        if db.day - dn.day > 0:
                            flag = "checked"
                            pyautogui.alert("Age is Less than 18 Years.")
                else:
                    flag = "checked"
                    pyautogui.alert("Age is Less than 18 Years.")
        return flag


    def get_save_data(self,event):
        if self.text57.GetValue():
            print("Photo Saved")
            self.save_file = os.path.join(IMAGE_PATH, "{}.jpg".format(self.text57.GetValue()))
            save_photo(self.text55.GetValue(),self.save_file)
        self.all_records = ["", self.chtext1.get(self.choice1[self.text1.GetSelection()]) if self.chtext1.get(
            self.choice1[self.text1.GetSelection()]) else None,
                            self.text2.GetValue(), self.text3.GetValue(), self.text4.GetValue(), self.text5.GetValue(),
                            self.text6.GetValue(), self.text7.GetValue(), self.text8.GetValue(), self.text9.GetValue(),
                            self.text10.GetValue(), self.text11.GetValue(), self.text12.GetValue(),
                            self.text13.GetValue(), self.text14.GetValue(), self.text15.GetValue(),
                            self.text16.GetValue(), self.text17.GetValue(), self.text18.GetValue(),
                            self.text19.GetValue(), self.text20.GetValue(), self.text21.GetValue(),
                            self.text22.GetValue(), self.text23.GetValue(), self.text24.GetValue(),
                            self.text25.GetValue(), self.text26.GetValue(), self.text27.GetValue(),
                            self.text28.GetValue(), self.text29.GetValue(), self.text30.GetValue(),
                            self.text31.GetValue(), self.text32.GetValue(), self.text33.GetValue(),
                            self.text34.GetValue(), self.text35.GetValue(), self.text36.GetValue(),
                            self.text37.GetValue(), self.text38.GetValue(), self.text39.GetValue(),
                            self.text40.GetValue(), self.text41.GetValue(), self.text42.GetValue(),
                            self.text43.GetValue(), self.text44.GetValue(), self.text45.GetValue(),
                            self.text46.GetValue(), self.text47.GetValue(), self.text48.GetValue(),
                            self.text49.GetValue(), self.text50.GetValue(), self.text51.GetValue(),
                            self.text52.GetValue(), self.text53.GetValue(), self.text54.GetValue(),
                            self.text55.GetValue(), self.text56.GetValue(), self.text57.GetValue(),
                            self.text58.GetValue(), self.text59.GetValue(), self.text60.GetValue(),
                            self.text61.GetValue(), self.text62.GetValue(),
                            F"{self.chtext63.get(self.choice63[self.text63.GetSelection()])}[{str(self.choice63[self.text63.GetSelection()])}]" if self.chtext63.get(
                                self.choice63[self.text63.GetSelection()]) else None,
                            F"{self.chtext64.get(self.choice64[self.text64.GetSelection()])}[{str(self.choice64[self.text64.GetSelection()])}]" if self.chtext64.get(
                                self.choice64[self.text64.GetSelection()]) else None,
                            F"{self.chtext65.get(self.choice65[self.text65.GetSelection()])}[{str(self.choice65[self.text65.GetSelection()])}]" if self.chtext65.get(
                                self.choice65[self.text65.GetSelection()]) else None,
                            F"{self.chtext66.get(self.choice66[self.text66.GetSelection()])}[{str(self.choice66[self.text66.GetSelection()])}]" if self.chtext66.get(
                                self.choice66[self.text66.GetSelection()]) else None,
                            self.text67.GetValue(),
                            F"{self.chtext68.get(self.choice68[self.text68.GetSelection()])}[{str(self.choice68[self.text68.GetSelection()])}]" if self.chtext68.get(
                                self.choice68[self.text68.GetSelection()]) else None,
                            F"{self.chtext69.get(self.choice69[self.text69.GetSelection()])}[{str(self.choice69[self.text69.GetSelection()])}]" if self.chtext69.get(
                                self.choice69[self.text69.GetSelection()]) else None,
                            F"{self.chtext70.get(self.choice70[self.text70.GetSelection()])}[{str(self.choice70[self.text70.GetSelection()])}]" if self.chtext70.get(
                                self.choice70[self.text70.GetSelection()]) else None,
                            F"{self.chtext71.get(self.choice71[self.text71.GetSelection()])}[{str(self.choice71[self.text71.GetSelection()])}]" if self.chtext71.get(
                                self.choice71[self.text71.GetSelection()]) else None,
                            F"{self.chtext72.get(self.choice72[self.text72.GetSelection()])}[{str(self.choice72[self.text72.GetSelection()])}]" if self.chtext72.get(
                                self.choice72[self.text72.GetSelection()]) else None,
                            self.text73.GetValue(), self.text74.GetValue(), self.text75.GetValue(),
                            self.text76.GetValue(), self.text77.GetValue(), self.text78.GetValue(),
                            self.text79.GetValue(), self.text80.GetValue(), self.text81.GetValue(),
                            self.text82.GetValue(),
                            F"{self.chtext83.get(self.choice83[self.text83.GetSelection()])}[{str(self.choice83[self.text83.GetSelection()])}]" if self.chtext83.get(
                                self.choice83[self.text83.GetSelection()]) else None,
                            F"{self.chtext84.get(self.choice84[self.text84.GetSelection()])}[{str(self.choice84[self.text84.GetSelection()])}]" if self.chtext84.get(
                                self.choice84[self.text84.GetSelection()]) else None,
                            F"{self.chtext85.get(self.choice85[self.text85.GetSelection()])}[{str(self.choice85[self.text85.GetSelection()])}]" if self.chtext85.get(
                                self.choice85[self.text85.GetSelection()]) else None,
                            self.text86.GetValue(), self.text87.GetValue(),
                            F"{self.chtext88.get(self.choice88[self.text88.GetSelection()])}[{str(self.choice88[self.text88.GetSelection()])}]" if self.chtext88.get(
                                self.choice88[self.text88.GetSelection()]) else None,
                            F"{self.chtext89.get(self.choice89[self.text89.GetSelection()])}[{str(self.choice89[self.text89.GetSelection()])}]" if self.chtext89.get(
                                self.choice89[self.text89.GetSelection()]) else None, self.text90.GetValue(),
                            self.text91.GetValue()]


        print(self.all_records[58])
        print(self.all_records[62])
        fl = self.check_records(self.all_records)

        if fl == "unchecked":
        # if True:
            finaldata = []
            finaldata.append(self.data["Timestamp"])
            finaldata.extend(self.all_records[1:17])
            finaldata.append(self.data["Permanent-Address"])
            finaldata.extend(self.all_records[17:37])
            finaldata.append(self.data['Choose One'])
            finaldata.extend(self.all_records[37:52])
            finaldata.append(self.data['Worked'])
            finaldata.extend(self.all_records[52:])
            result = save_data(finaldata,self.ind+2)
            if result == "saved":
                pyautogui.alert("Data Saved")
            else:
                pyautogui.alert("Unable to  save data")
            if self.text55.GetValue() and self.text56.GetValue() and self.text57.GetValue() and self.text58.GetValue() and self.text62.GetValue() and self.chtext63.get(self.choice63[self.text63.GetSelection()]) and self.chtext64.get(self.choice64[self.text64.GetSelection()]) and self.chtext65.get(self.choice65[self.text65.GetSelection()]) and self.chtext66.get(self.choice66[self.text66.GetSelection()]) and self.text67.GetValue() and self.chtext68.get(self.choice68[self.text68.GetSelection()]) and self.chtext69.get(self.choice69[self.text69.GetSelection()]) and self.chtext70.get(self.choice70[self.text70.GetSelection()]) and self.chtext71.get(self.choice71[self.text71.GetSelection()]) and self.chtext72.get(self.choice72[self.text72.GetSelection()]) and self.text73.GetValue() and self.text74.GetValue() and self.text76.GetValue() and self.text77.GetValue() and self.text78.GetValue() and self.text79.GetValue() and self.text80.GetValue() and self.text81.GetValue() and self.chtext83.get(self.choice83[self.text83.GetSelection()]) and self.chtext84.get(self.choice84[self.text84.GetSelection()]) and self.chtext85.get(self.choice85[self.text85.GetSelection()]) and self.text87.GetValue() and self.chtext88.get(self.choice88[self.text88.GetSelection()]) and self.chtext89.get(self.choice89[self.text89.GetSelection()]) and self.text90.GetValue() and self.text91.GetValue():
                self.sendmail_button.Enable()


    def mail_send(self,event):
        if self.text55.GetValue() and self.text56.GetValue() and self.text57.GetValue() and self.text58.GetValue() and self.text62.GetValue() and self.chtext63.get(self.choice63[self.text63.GetSelection()]) and self.chtext64.get(self.choice64[self.text64.GetSelection()]) and self.chtext65.get(self.choice65[self.text65.GetSelection()]) and self.chtext66.get(self.choice66[self.text66.GetSelection()]) and self.text67.GetValue() and self.chtext68.get(self.choice68[self.text68.GetSelection()]) and self.chtext69.get(self.choice69[self.text69.GetSelection()]) and self.chtext70.get(self.choice70[self.text70.GetSelection()]) and self.chtext71.get(self.choice71[self.text71.GetSelection()]) and self.chtext72.get(self.choice72[self.text72.GetSelection()]) and self.text73.GetValue() and self.text74.GetValue() and self.text76.GetValue() and self.text77.GetValue() and self.text78.GetValue() and self.text79.GetValue() and self.text80.GetValue() and self.text81.GetValue() and self.chtext83.get(self.choice83[self.text83.GetSelection()]) and self.chtext84.get(self.choice84[self.text84.GetSelection()]) and self.chtext85.get(self.choice85[self.text85.GetSelection()]) and self.text87.GetValue() and self.chtext88.get(self.choice88[self.text88.GetSelection()]) and self.chtext89.get(self.choice89[self.text89.GetSelection()]) and self.text90.GetValue() and self.text91.GetValue():
            try:
                response = update_series_data(list(self.temp.values()))
                print(response)
            except AttributeError as e:
                print("Error: ", e)
            empno = self.all_records[57]
            dob = self.all_records[28]
            dob = dob[3:5] + '/' + dob[0:2] + '/' + dob[-4:]
            doj = self.all_records[58]
            doj = doj[3:5] + '/' + doj[0:2] + '/' + doj[-4:]
            epf_dict = {'Display Name': self.all_records[2]+" "+self.all_records[3]+" "+self.all_records[4],
                        'Fathers/Husband Name': self.all_records[6],
                        'Date Of Birth': dob,
                        'Gender -M/F': self.all_records[5],
                        'Marital Status': self.all_records[9],
                        'Email ID': self.all_records[35],
                        'Personal Mobile No': self.all_records[36],
                        'UAN': self.all_records[46],
                        'P.F. A/c No.(6 Digits)': "",
                        'Primary Bank A/c No': self.all_records[24],
                        'Primary IFSC': self.all_records[23],
                        'Aadhaar Card No': self.all_records[33],
                        'Permanent A/c No': self.all_records[24],
                        'Date Of Joining': self.all_records[58],
                        'Name':self.all_records[2]+" "+self.all_records[3]}
            tmpl_path = os.path.join(UTIL_PATH,"template.pdf")
            tmpl_save = os.path.join(TMPL_PATH,f"EPF {self.all_records[57]}.pdf")
            create_epf(epf_dict,tmpl_path,tmpl_save)
            lst = [63, 64, 65, 66, 68, 69, 70, 71, 72, 83, 84, 85, 88, 89]
            excel_data = [v.split("[")[1].split("]")[0] if i in lst else v for i, v in enumerate(self.all_records)]

            create_json(JSON_PATH,empno,[self.all_records[58],self.all_records[35],self.all_records[2]])

            staffing_name = f"{empno}({self.all_records[2]} {self.all_records[4]}).xlsx"
            create_id_staffing(excel_data,os.path.join(STAFFING_PATH,staffing_name))

            core_name = f"{empno}({self.all_records[2]} {self.all_records[4]}).xlsx"
            create_id_core(excel_data,os.path.join(CORE_PATH,core_name))

            idcard_name = f"{empno}({self.all_records[2]} {self.all_records[4]}).xlsx"
            create_id_card(excel_data,os.path.join(IDCARD_PATH,idcard_name),self.chtext70[excel_data[70]])

            ascent_name = "Ascent({}).xlsx".format(get_date())
            create_ascent(excel_data,os.path.join(ASCENT_PATH,ascent_name))

            zing_name = "Zing.xlsx"
            create_zing(excel_data,os.path.join(ZING_PATH,zing_name))

            full_name = "all_records.xlsx"
            full_record(excel_data,os.path.join(FULLDATA_PATH,full_name))

            empno = self.all_records[57]
            create_mis(os.path.join(MIS_PATH,"MIS_File.xlsx"),empno)
            ltr_name = os.path.join(WORD_PATH,F"{empno}.docx")

            key = None
            if str(empno)[0] == "S":
                if self.all_records[64].split("[")[0] == "CONTRACT":
                    key = "On_rolls"
                else:
                    key = "On_Contract"
            elif str(empno)[0] == "9":
                pass
            if  self.all_records[66].split("[")[0][0:1] == "L" and self.all_records[64].split("[")[0] == "GENERAL":
                key = "L_Band"
            elif self.all_records[66].split("[")[0][0:2] == "MB" and self.all_records[64].split("[")[0] == "GENERAL":
                key = "MB_and_Above"
            elif self.all_records[66].split("[")[0][0:2] == "MT" and self.all_records[64].split("[")[0] == "GENERAL":
                key = "MT"
            elif self.all_records[66].split("[")[0][0:2] == "MS" and self.all_records[66].split("[")[0][2] != "0" and self.all_records[64].split("[")[0] == "GENERAL":
                key = "MS0_MS4"
            elif self.all_records[66].split("[")[0][0:2] == "MS" and self.all_records[66].split("[")[0][2] != "0" and self.all_records[64].split("[")[0] == "CONTRACT":
                key = "Contract_MS0_MS4"
            elif self.all_records[66].split("[")[0][0:3] == "MS0" and self.all_records[64].split("[")[0] == "CONTRACT":
                key = "Trainee"

            if key:
                get_letter(excel_data, key, ltr_name)
            delete_record(self.all_records)
            self.Destroy()
            rec = [self.all_records[2], self.all_records[4], self.all_records[67], self.all_records[84].split("[")[0],
                   self.all_records[70].split("[")[0],
                   f"{self.all_records[58][3:5]}/{self.all_records[58][:2]}/{self.all_records[58][6:]}"]

            tt = []
            tt.append(self.all_records[91])
            tmplist = [[os.path.join(IDCARD_PATH, idcard_name),self.save_file],[os.path.join(STAFFING_PATH,staffing_name)],[os.path.join(CORE_PATH,core_name)],[]]
            ##################################

            body = """Hello there,
		Please Find the attachment for {} Excel"""
            confbody = f"""

Dear Sir/ Madam,



We take immense pleasure to introduce (Name- {rec[0]} {rec[1]}), who has joined us as {rec[2]}. She/He will be operating from {rec[4]}.



Details are as follows

Name: {rec[0]}

Designation: {rec[2]}

Department: {rec[3]}

Location: {rec[4]}

Date: {rec[5]}




Regards,

HR Team."""
            mldf = pandas.read_excel(mailing_excel)
            mlvl = mldf.values
            data_mail = []
            for i,line in enumerate(mlvl):
                if i < 3:
                    body = body.format(line[3])
                    tmp = [str(line[1]).split(","), str(line[2]).split(","), line[3],body,tmplist[i]]
                else:
                    body = confbody
                    tmp = [tt, [], line[3],body,[]]
                data_mail.append(tmp)

            ##################################

            otf = os.path.join(APPN_PATH, f"{os.path.basename(ltr_name).split('.docx')[0]}.pdf")
            convert_to_pdf(ltr_name, otf)
            for md in data_mail:
                sendoutlookmail(md[0],md[1],md[2],md[3],md[4],)
            # send_joining_confirmation(jcd, tt)
            # send_mails("ID Card", os.path.join(IDCARD_PATH, idcard_name),self.save_file)
            # send_mails("Email ID Staffing form", os.path.join(STAFFING_PATH,staffing_name))
            # send_mails("Email ID Core form", os.path.join(CORE_PATH,core_name))
            # send_zing(os.path.join(CHRM_PATH,"chromedriver.exe"))
        else:
            pyautogui.alert("Please Fill all the Entries to continue.\n\n\nPress OK to Abort.")


if __name__ == '__main__':
    STAFFING_PATH = os.path.join(os.getcwd(),"Files","Staffing")
    CORE_PATH = os.path.join(os.getcwd(),"Files","core")
    IDCARD_PATH = os.path.join(os.getcwd(),"Files","IDCard")
    WORD_PATH = os.path.join(os.getcwd(),"Files","Word Files")
    APPN_PATH = os.path.join(os.getcwd(),"Files","Appointment letter Pdf")
    ASCENT_PATH = os.path.join(os.getcwd(),"Files","Ascent")
    ZING_PATH = os.path.join(os.getcwd(),"Files","Zing")
    FULLDATA_PATH = os.path.join(os.getcwd(),"Files","All Records")
    UTIL_PATH = os.path.join(os.getcwd(),"Files","util")
    TMPL_PATH = os.path.join(os.getcwd(),"Files","EPF")
    IMAGE_PATH = os.path.join(os.getcwd(),"Files","images")
    JSON_PATH = os.path.join(os.getcwd(),"Files","Json")
    MIS_PATH = os.path.join(os.getcwd(),"Files","Mis")
    CHRM_PATH = os.path.join(os.getcwd(),"Files","chromedriver")
    mailing_excel = os.path.join(UTIL_PATH,"mailing_details.xlsx")
    make_dir(STAFFING_PATH,CORE_PATH, IDCARD_PATH,ASCENT_PATH,APPN_PATH,ZING_PATH,FULLDATA_PATH,UTIL_PATH,IMAGE_PATH,WORD_PATH,JSON_PATH,MIS_PATH,TMPL_PATH,CHRM_PATH)
    app = wx.App()
    frame = Checker_Process(parent=None, id=-1)
    frame.Show()
    app.MainLoop()