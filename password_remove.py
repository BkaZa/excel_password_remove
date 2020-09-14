import os, sys
import glob
import win32com.client

cur_dir = os.getcwd()

def remove_password(fn, pwd):
    full_path = cur_dir+"\\"+fn
    
    xcl = win32com.client.Dispatch("Excel.Application")
    xcl.DisplayAlerts = False
    wb = xcl.Workbooks.Open(full_path, False, False, None, pwd)
    
    wb.Unprotect(pwd)
    wb.UnprotectSharing(pwd)

    sv = save_file(fn)
    
    wb.SaveAs(sv, None, '', '')
    xcl.Quit()

    return sv

def save_file(fn):
    unlock_path = cur_dir+"\\unlock\\"
    if not os.path.exists(unlock_path):
        os.makedirs(unlock_path)
        
    sv = r""+unlock_path+fn
    
    return sv

input_pwd = input("Input Password : ")

file_list = os.listdir()
file_ext = tuple([".xls", ".xlsx"])

total=1
for index, file_name in enumerate(file_list):
    if file_name.endswith(file_ext):
        try:
            unlock = remove_password(file_name, input_pwd)
            print("Unlock To: "+unlock)
            total = (total+1)
        except Exception as error:
            print("Check Your Password Agian!!!")
            total = 0
            break

print("\nUnlock Success Total: "+str(total))
input("Enter key for exit...")
