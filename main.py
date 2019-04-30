""" FSS NOC Tool Installer """
#
#   Author: Virgil Hoover
#   License found in '../Documentation/License.txt'

import tkinter as tk
from os import getenv, makedirs
from os.path import isdir, isfile
from re import sub
from shutil import copy, copytree, move, rmtree
from subprocess import Popen, check_output
from sys import version
from time import sleep, time
from tkinter import filedialog
from tkinter.messagebox import showinfo
from tkinter.ttk import OptionMenu
from urllib import request
from webbrowser import open_new

SERVER = input('What is the server URL? ')
DESTINATION = 'C:\\Scripts'
COMPRESSION = 'C:\\Program Files\\7-Zip\\7z.exe'
CONNECTION = 'C:\\Windows\\System32\\PsExec.exe'


def about():
    """ Information about the program. """
    showinfo('FSS NOC Tools', 'Author: Virgil Hoover\n'
                              'Created: February 20, 2019\n'
                              'Modified: April 6, 2019\n'
                              'Version: 2.0')


def email():
    """ Contact the creator. """
    exists = getenv('UserProfile') + '\\ApppData\\Local\\Microsoft\\Outlook\\*.ost'
    if exists:
        open_new(r'virgilhoover@gmail.com')
    else:
        open_new('https://mail.office365.com/')
    return


def remote_connection():
    if not CONNECTION:
        download('https://download.sysinternals.com/files/PSTools.zip', '\\PSTools.zip')
        unzip('PSTools.zip', 'PSTools')
        copy(DESTINATION + '\\PsTools\\PsExec.exe', 'C:\\Windows\\System32\\PsExec.exe')
    else:
        return


def check_powershell():
    p = check_output(r'reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine /v '
                     r'PowerShellVersion')
    v = p.split()
    pwr_reqs = str(v[3])
    pwr_reqs = sub("[b\\']", '', pwr_reqs)
    if not pwr_reqs >= '5.0.0':
        download('https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/'
                 'Win7AndW2K8R2-KB3191566-x64.zip', '\\PowerShell5.1.zip')
        unzip('PowerShell5.1.zip', 'SnIPT')
        execute_installer(DESTINATION + '\\SnIPT\\Win7AndW2K8R2-KB3191566-x64.msu')


def download(url, file):
    """ Obtain the request file from the server. """
    exists = isdir(DESTINATION)
    if not exists:
        makedirs(DESTINATION)
    request.urlretrieve(url, DESTINATION + file)
    return


def tool_run(file):
    """ Execute the requested tool. """
    Popen('cmd /c ' + file)
    return


def archive(file):
    """ archive the Tool zip file for later. """
    must_end = time() + 3
    exists = isdir(DESTINATION + '\\Archive')
    if not exists:
        makedirs(DESTINATION + '\\Archive')
    if time() > must_end:
        move(DESTINATION + file, DESTINATION + '\\Archive\\' + file)
    else:
        sleep(3)
        move(DESTINATION + file, DESTINATION + '\\Archive\\' + file)
    return


def delete():
    """ Delete the current list of tools installed. """
    folder = filedialog.askdirectory()
    exists = isdir(folder)
    if exists:
        rmtree(folder)
    else:
        return


def unzip(file, folder):
    if not COMPRESSION:
        download('https://www.7-ziorg/a/7z1806-x64.exe', '\\7zip.exe')
        Popen(DESTINATION + '\\7zip.exe')
    Popen(COMPRESSION + ' e ' + DESTINATION + file + ' -o' + DESTINATION + '\\' + folder)
    archive(file)


def execute_installer(file):
    Popen(DESTINATION + file)
    return


def rlpi_install():
    """ Install the RLPI Tool. """
    file = '\\RLPI.zip'
    url = SERVER + '/RL/RLPI.zip'
    download(url, file)
    sql_req = 'C:\\Program Files (x86)\\Microsoft SQL w.SERVER\\Client SDK\\ODBC\\130\\Tools\\Binn\\SQLCMD.EXE'
    if not sql_req:
        download('https://go.microsoft.com/fwlink/?linkid=2043518', '\\MsSqlCmdLnUtils.msi')
        execute_installer('MsSqlCmdLnUtils.msi')
    remote_connection()
    unzip(file, 'RLPI')
    return


def snipt_install():
    """ Install the SnIPT Tool. """
    file = '\\SnIPT.zip'
    url = SERVER + '/SIP/SnIPT.zip'
    download(url, file)
    nmap_req = 'C:\\Program Files (x86)\\Nmap\\nmap.exe'
    check_powershell()
    remote_connection()
    if not nmap_req:
        download('https://nmaorg/dist/nmap-7.70-setuexe', '\\Nmap.exe')
        execute_installer(DESTINATION + '\\Nmap.exe')
    unzip(file, 'SnIPT')
    return


def ups_install():
    """ Install the UPS Tool. """
    file = '\\UPS.zip'
    url = SERVER + '/UPS/UPS.zip'
    download(url, file)
    check_powershell()
    unzip(file, 'UPS')
    return


def time_install():
    """ Install the Time Conversion Tool. """
    file = '\\timeconv.zip'
    url = SERVER + '/Time/timeconv.zip'
    download(url, file)
    unzip(file, 'TimeConvert')
    return


def sap_install():
    """ Install the SAP Tool. """
    file = '\\SAP.zip'
    url = SERVER + '/SAP/SAP.zip'
    download(url, file)
    unzip(file, 'SAP')
    makedirs('C:\\Installs')
    sleep(3)
    copytree(DESTINATION + '\\SAP', 'C:\\Installs\\SAP')
    sleep(3)
    return


def gpupdate_install():
    """ Install the Update Tool. """
    file = '\\gpupdate.zip'
    url = SERVER + '/GPU/gpupdate.zip'
    download(url, file)
    unzip(file, 'GPUpdate')
    return


def legal_install():
    """ Install the Legal Tool. """
    file = '\\LegalHoldTool.zip'
    url = SERVER + '/LEGAL/LegalHoldTool.zip'
    download(url, file)
    py_req = version
    if not py_req >= '3.6.5':
        download('https://www.python.org/ftp/python/3.7.2/python-3.7.2.exe', '\\python-3.7.2.exe')
        execute_installer(DESTINATION + '\\python-3.7.2.exe')
    unzip(file, 'Legal')
    return


def logon():
    """ Install the Last Logon Tool. """
    file = '\\RLLastLogon.zip'
    url = SERVER + '/RL/RLLastLogon.zip'
    download(url, file)
    check_powershell()
    unzip(file, 'LastLogon')
    return


def printer():
    """ Install the Printer Tool. """
    file = '\\PrinterTool.exe'
    url = SERVER + 'PT/PrinterTool.exe'
    download(url, file)
    makedirs(DESTINATION + '\\PrinterTool')
    copy(file, DESTINATION + '\\PrinterTool')
    archive(file)
    return


def hostname():
    """ Install the Host Tool. """
    file = '\\HostnameConverter.zip'
    url = SERVER + '/HOST/HostnameConverter.zip'
    download(url, file)
    unzip(file, 'Hostname')
    return


def transfer():
    """ Install the File transfer Tool. """
    file = '\\FSSDataTransfer.zip'
    url = SERVER + '/DT/FSSDataTransfer.zip'
    download(url, file)
    unzip(file, 'Transfer')
    return


def time_change():
    """ Install the Timezone Change Tool. """
    file = '\\timeChange.zip'
    url = SERVER + '/Time/timeChange.zip'
    download(url, file)
    remote_connection()
    unzip(file, 'TimeChange')
    return


class WindowStructure(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)

        container.pack(side='top', fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, RunWindow):
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky='nsew')

        """ This class is the main gusto of the program. """
        # Main Window
        self.title('NOC Tools')
        window_width = self.winfo_reqwidth()
        window_height = self.winfo_reqheight()
        self.resizable(False, False)
        self.focusmodel('active')
        position_right = int(self.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(self.winfo_screenheight() / 3 - window_height / 3)
        self.geometry("+{}+{}".format(position_right, position_down))

        self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        # Menu bar
        self.nav_button1 = tk.Button(self, text='Execute', command=lambda: controller.show_frame(RunWindow))
        self.nav_button1.grid(row=1, column=1, sticky='ew')
        self.nav_button2 = tk.Button(self, text='Remove', command=delete)
        self.nav_button2.grid(row=1, column=2, sticky='ew')
        self.nav_button3 = tk.Button(self, text='Email', command=email)
        self.nav_button3.grid(row=1, column=3, sticky='ew')
        self.nav_button4 = tk.Button(self, text='About', command=about)
        self.nav_button4.grid(row=1, column=4, sticky='ew')

        # Body of Install Window
        self.description_label = tk.Label(self, text='FSS NOC Tool Installer', font='Helvetica 18 bold',
                                          anchor='center', background='#0033A0', foreground='white', relief='sunken')
        self.description_label.grid(row=2, column=1, columnspan=4, sticky='ew')
        self.tools = [' Make your selection ',
                      'Rounding Laptop Tool',
                      'SnIPT Tool',
                      'UPS Generator',
                      'Timezone Conversion',
                      'SAP Tool',
                      'Remote GPUpdate',
                      'Legal Hold Tool',
                      'RL Last Logon',
                      'Printer Tool',
                      'Hostname Converter',
                      'Data Transfer Tool',
                      'Timezone Changer']
        self.tool_label = tk.StringVar(self)
        self.opt_menu = OptionMenu(self, self.tool_label, *sorted(self.tools))
        self.opt_menu.grid(row=4, column=1, columnspan=4, sticky='ew', pady=20)
        self.tool_label.trace('w', self.tool_selection)
        self.start = 'This tool will install all the selected programs to your local system. ' \
                     'Please select the item you need and then click Install.\n\n ' \
                     'Please provide any issues, or requests to virgil.hoover@fmc-na.com'
        self.tool_description = tk.Label(self, text=self.start, foreground='grey', relief='sunken',
                                         wraplength=375, justify='left')
        self.tool_description.grid(row=5, column=1, columnspan=4, sticky='ew', padx=10, pady=5)

        # Action Buttons
        self.tool = tk.Button(self, text='Install', background='green', height=4, width=10, command=self.tool_install)
        self.tool.grid(row=6, column=1, columnspan=2, padx=10, pady=20)
        self.cancel = tk.Button(self, text='Cancel', background='red', height=4, width=10, command=self.quit)
        self.cancel.grid(row=6, column=3, columnspan=2, padx=10, pady=20)

        # Status Bar
        self.progress = tk.Label(self, text='', relief='sunken', anchor='w', )
        self.progress.grid(row=7, column=1, columnspan=4, sticky='sew', pady=5)

    def tool_selection(self, *args):
        """ Sets the tool_description field based off the tool selected. """
        self.tool_label.get()
        if self.tool_label.get() == 'Rounding Laptop Tool':
            self.tool_description.configure(text='This tool installs the clinic Lexmark, '
                                                 'sets the timezone and wipes the last '
                                                 'logged in user.\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - SQL Commandline Utilities'
                                                 '\n - 7zip \n - PsExec')
        elif self.tool_label.get() == 'SnIPT Tool':
            self.tool_description.configure(text='The SnIPT tool allows you to scan entire '
                                                 'subnet to find all printers that are '
                                                 'installed on all live PCs at the time of '
                                                 'execution.\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip '
                                                 '\n - PowerShell v5.0 or higher'
                                                 '\n - PsExec \n - Nmap v7.7')
        elif self.tool_label.get() == 'UPS Generator':
            self.tool_description.configure(text='This is a PowerShell script that pulls '
                                                 'the Remedy information out of all the '
                                                 'outstanding UPS tickets.\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip '
                                                 '\n - PowerShell v5.0 or higher')
        elif self.tool_label.get() == 'Timezone Conversion':
            self.tool_description.configure(text='This tool tells you at what time a '
                                                 'maintenance window should be set for.'
                                                 '\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip')
        elif self.tool_label.get() == 'SAP Tool':
            self.tool_description.configure(text='This tool remotely installs SAP 7.40 on '
                                                 'a windows computer.\n\nThis tool and the '
                                                 'following dependencies will be installed: '
                                                 '\n - 7zip')
        elif self.tool_label.get() == 'Remote GPUpdate':
            self.tool_description.configure(text='This tool is used to to run gpupdate '
                                                 'on a remote computer, useful before '
                                                 'running the Rounding Laptop tool.'
                                                 '\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip')
        elif self.tool_label.get() == 'Legal Hold Tool':
            self.tool_description.configure(text='The Legal Hold Tool allows users to get the '
                                                 'list of all users of a machine and transfer user '
                                                 'files to East or West division legal directories.'
                                                 '\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip'
                                                 '\n - Python 3.6.5 or higher')
        elif self.tool_label.get() == 'RL Last Logon':
            self.tool_description.configure(text='A tool that scans for every service tag in the '
                                                 '"RL" queue to show the rounding laptops that have '
                                                 'been logged in since configured at the NOC, also '
                                                 'grabbing the most recent I\n\nThis tool and the '
                                                 'following dependencies will be installed: \n - 7zip '
                                                 '\n - PowerShell v5.0 or higher')
        elif self.tool_label.get() == 'Printer Tool':
            self.tool_description.configure(text='This tool allows for remote printer management '
                                                 'of a windows computer, including installing, '
                                                 'removing printers, and renaming printers')
        elif self.tool_label.get() == 'Hostname Converter':
            self.tool_description.configure(text='The Hostname converter tool allows the '
                                                 'user to get the IP address of a host '
                                                 'automatically.\n\nThis tool and the following '
                                                 'dependencies will be installed: \n - 7zip')
        elif self.tool_label.get() == 'Data Transfer Tool':
            self.tool_description.configure(text='FSS Data Transfer Tool allows users '
                                                 'to transfer files from an old '
                                                 'computer to a new computer.\n\nThis '
                                                 'tool and the following dependencies '
                                                 'will be installed: \n - 7zip')
        elif self.tool_label.get() == 'Timezone Changer':
            self.tool_description.configure(text='This tool sets the timezone on a remote '
                                                 'windows computer.\n\nThis tool and the '
                                                 'following dependencies will be installed: '
                                                 '\n - 7zip \n - PowerShell v5.0 or higher'
                                                 '\n - PsExec')
        else:
            self.tool_description.configure(text='Something went wrong.')
        return self.tool_description

    def tool_install(self, *args):
        """ Installs the selected tool"""
        self.tool_label.get()
        if self.tool_label.get() == 'Rounding Laptop Tool':
            rlpi_install()
            self.progress.configure(text='RLPI files installed')
        elif self.tool_label.get() == 'SnIPT Tool':
            snipt_install()
            self.progress.configure(text='SnIPT files installed')
        elif self.tool_label.get() == 'UPS Generator':
            ups_install()
            self.progress.configure(text='UPS Generator files installed')
        elif self.tool_label.get() == 'Timezone Conversion':
            time_install()
            self.progress.configure(text='Timezone Conversion files installed')
        elif self.tool_label.get() == 'SAP Tool':
            sap_install()
            self.progress.configure(text='SAP files installed')
        elif self.tool_label.get() == 'Remote GPUpdate':
            gpupdate_install()
            self.progress.configure(text='GPUpdate files installed')
        elif self.tool_label.get() == 'Legal Hold Tool':
            legal_install()
            self.progress.configure(text='Legal Hold files installed')
        elif self.tool_label.get() == 'RL Last Logon':
            logon()
            self.progress.configure(text='RL Last Logon files installed')
        elif self.tool_label.get() == 'FSS Printer Tool':
            printer()
            self.progress.configure(text='Printer Tool files installed')
        elif self.tool_label.get() == 'Hostname Converter':
            hostname()
            self.progress.configure(text='Hostname files installed')
        elif self.tool_label.get() == 'FSS Data Transfer Tool':
            transfer()
            self.progress.configure(text='Data Transfer files installed')
        elif self.tool_label.get() == 'Timezone Changer':
            time_change()
            self.progress.configure(text='Timezone Changer files installed')
        else:
            self.tool_description.configure(text='Something went wrong!!!')
        self.tool_description.configure(text=self.start)
        return self.progress


class RunWindow(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        """ This is the second frame, the one that gives the user the ability to execute the installed tool. """

        # Menu bar
        nav_button1 = tk.Button(self, text='Install', command=lambda: controller.show_frame(StartPage))
        nav_button1.grid(row=1, column=1, sticky='ew')
        nav_button2 = tk.Button(self, text='Remove', command=delete)
        nav_button2.grid(row=1, column=2, sticky='ew')
        nav_button3 = tk.Button(self, text='Email', command=email)
        nav_button3.grid(row=1, column=3, sticky='ew')
        nav_button4 = tk.Button(self, text='About', command=about)
        nav_button4.grid(row=1, column=4, sticky='ew')

        # Body of run window
        title_label = tk.Label(self, text='FSS NOC Tool Executor', font='Helvetica 18 bold', anchor='center',
                               background='#0033A0', foreground='white', relief='sunken', width=25)
        title_label.grid(row=2, column=1, columnspan=4, sticky='ew')
        disclaimer = tk.Label(self, text='Note buttons will be disabled for all tools not installed.')
        disclaimer.grid(row=3, column=1, columnspan=4, sticky='ew', pady=5)

        # Button 1
        transfer_file = DESTINATION + '\\Transfer\\FSS_Data_Transfer_Tool_v1.exe'
        data_transfer = tk.Button(self, text='FSS Data Transfer Tool', command=lambda: tool_run(transfer_file),
                                  anchor='w', font='Helvetica 10 bold')
        data_transfer.grid(row=4, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(transfer_file):
            data_transfer.configure(state='disabled', font='Helvetica 10')

        # Button 2
        printer_file = DESTINATION + '\\PrinterTool\\PrinterTool.exe'
        printer_tool = tk.Button(self, text='FSS Printer Tool', command=lambda: tool_run(printer_file),
                                 anchor='w', font='Helvetica 10 bold')
        printer_tool.grid(row=4, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(printer_file):
            printer_tool.configure(state='disabled', font='Helvetica 10')

        # Button 3
        hostname_file = DESTINATION + '\\Hostname\\HostnameConverterv1.0.bat'
        hostname_conv = tk.Button(self, text='Hostname Converter', command=lambda: tool_run(hostname_file),
                                  anchor='w', font='Helvetica 10 bold')
        hostname_conv.grid(row=5, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(hostname_file):
            hostname_conv.configure(state='disabled', font='Helvetica 10')

        # Button 4
        legal_file = DESTINATION + '\\Legal\\LegalHold_Tool_v2.5.bat'
        legal = tk.Button(self, text='Legal Hold Tool', command=lambda: tool_run(legal_file),
                          anchor='w', font='Helvetica 10 bold')
        legal.grid(row=5, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(legal_file):
            legal.configure(state='disabled', font='Helvetica 10')

        # Button 5
        rl_file = DESTINATION + '\\LastLogon\\RLLastLogon.bat'
        rl_logon = tk.Button(self, text='RL Last Logon Tool', command=lambda: tool_run(rl_file),
                             anchor='w', font='Helvetica 10 bold')
        rl_logon.grid(row=6, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(rl_file):
            rl_logon.configure(state='disabled', font='Helvetica 10')

        # Button 6
        update_file = DESTINATION + '\\Update\\gpupdate.bat'
        gpu = tk.Button(self, text='Remote GPUpdate', command=lambda: tool_run(update_file),
                        anchor='w', font='Helvetica 10 bold')
        gpu.grid(row=6, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(update_file):
            gpu.configure(state='disabled', font='Helvetica 10')

        # Button 7
        rlpi_file = DESTINATION + '\\RLPI\\RLPI.bat'
        rlpi = tk.Button(self, text='Rounding Laptop Tool', command=lambda: tool_run(rlpi_file),
                         anchor='w', font='Helvetica 10 bold')
        rlpi.grid(row=7, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(rlpi_file):
            rlpi.configure(state='disabled', font='Helvetica 10')

        # Button 8
        sap_file = DESTINATION + '\\SAP\\SAPInstall.bat'
        sap = tk.Button(self, text='SAP Remote Install', command=lambda: tool_run(sap_file),
                        anchor='w', font='Helvetica 10 bold')
        sap.grid(row=7, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(sap_file):
            sap.configure(state='disabled', font='Helvetica 10')

        # Button 9
        snipt_file = DESTINATION + '\\SnIPT\\SnIPTv.2.bat'
        snipt = tk.Button(self, text='SnIPT', command=lambda: tool_run(snipt_file),
                          anchor='w', font='Helvetica 10 bold')
        snipt.grid(row=8, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(snipt_file):
            snipt.configure(state='disabled', font='Helvetica 10')

        # Button 10
        time_conv = DESTINATION + '\\TimeConvert\\timeconv.bat'
        lcl_time = tk.Button(self, text='Time Conversion Tool', command=lambda: tool_run(time_conv),
                             anchor='w', font='Helvetica 10 bold')
        lcl_time.grid(row=8, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(time_conv):
            lcl_time.configure(state='disabled', font='Helvetica 10')

        # Button 11
        ups_file = DESTINATION + '\\UPS\\UPSGenerator.bat'
        ups = tk.Button(self, text='UPS Generator', command=lambda: tool_run(ups_file),
                        anchor='w', font='Helvetica 10 bold')
        ups.grid(row=9, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(ups_file):
            ups.configure(state='disabled', font='Helvetica 10')

        # Button 12
        change_file = DESTINATION + '\\TimeChange\\timeChanger.bat'
        change = tk.Button(self, text='Timezone Changer', command=lambda: tool_run(change_file),
                           anchor='w', font='Helvetica 10 bold')
        change.grid(row=9, column=3, columnspan=2, sticky='ew', padx=5, pady=5)
        if not isfile(change_file):
            change.configure(state='disabled', font='Helvetica 10')

        # Status Bar
        status = tk.Label(self, text='', anchor='w', relief='sunken')
        status.grid(row=10, column=1, columnspan=4, sticky='ew', pady=5)


if __name__ == '__main__':
    app = WindowStructure()
    app.mainloop()
