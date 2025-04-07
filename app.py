import pandas as pd
import openpyxl
import re, glob, os
import datetime as datetime

import tkinter as tk #get user input

def get_username():
    root = tk.Tk()
    root.title("Esker Invoice Download")
    root.geometry("380x230")

    name_var = tk.StringVar()
    username=None

    def submit_username():
        nonlocal username
        username = name_var.get()
        print("Username:", username)  # Print username within the function
        root.destroy()  # Close the window after submission
        #return name  # Return the username

    name_label = tk.Label(root, text='Username', font=('calibre', 10, 'bold'))
    name_entry = tk.Entry(root, textvariable=name_var, font=('calibre', 10, 'normal'))

    submit_btn = tk.Button(root, text='Submit', command=submit_username)

    name_label.grid(row=0, column=0, padx=5, pady=5)
    name_entry.grid(row=0, column=1, padx=5, pady=5)
    submit_btn.grid(row=1, column=1, padx=5, pady=5)
    root.mainloop()  # Start the Tkinter event loop and store the returned value
    return username  # Return the username

global username
username=get_username()  # Call the function to get the username


def main():

    path_invoice_name_number = r'C:/Users/'+username+'/Downloads/esker_merged/'
    df_invoice_name_number = pd.read_excel(path_invoice_name_number+'invoice_rename.xlsx', sheet_name='invoice_rename', engine='openpyxl')

    list_invoice_name = df_invoice_name_number['invoice_name'].tolist()
    list_invoice_numbr = df_invoice_name_number['invoice'].tolist()
    list_posting_date = df_invoice_name_number['posting_date'].tolist()
    list_invoice_numbr = [str(i) for i in list_invoice_numbr] # convert list_invoice_numbr elements to string
    list_posting_date = [str(i) for i in list_posting_date]
    try:

        list_posting_date = [x.replace(" 00:00:00", "") for x in list_posting_date]
    except Exception as e:
        print(e)
    try:
        list_posting_date = [pd.to_datetime(x).strftime('%d%m%Y') for x in list_posting_date]
    except Exception as e:
        print(e)

    list_of_lists = [list_invoice_numbr, list_invoice_name]
    list_of_lists_2 = [list_invoice_numbr, list_posting_date]
    def get_dict_invoice_name_number(list_of_lists):
        return {str(z[0]): (z[1]) for z in zip(*list_of_lists)}

    def get_dict_invoice_name_posting_date(list_of_lists_2):
        return {str(z[0]): (z[1]) for z in zip(*list_of_lists_2)}

    dict_invoice_name_number = get_dict_invoice_name_number(list_of_lists)
    dict_invoice_name_posting_date = get_dict_invoice_name_posting_date(list_of_lists_2)

    def format_dictionary_keys_in_place(dictionary_invoice):
        # Iterate over a copy of the keys to avoid RuntimeError
        for key in list(dictionary_invoice.keys()):
            new_key = key.replace(':', '_').replace('/', '_')
            if new_key != key:  # Only update if the key has changed
                dictionary_invoice[new_key] = dictionary_invoice.pop(key)
        return dictionary_invoice  # Return the modified dictionary

    format_dict_invoice_name_number = format_dictionary_keys_in_place(dict_invoice_name_number)
    format_dict_invoice_name_posting_date = format_dictionary_keys_in_place(dict_invoice_name_posting_date)
    #print(format_dict_invoice_name_number)
    #print(format_dict_invoice_name_posting_date)

    def convert_and_append_merged_safe(list_invoice_number):
        new_list = []
        for item in list_invoice_number:
            if item is not None:  # Skip None values (or add other checks)
                try:
                    item = item.replace('/', '_').replace(':', '_')
                    new_list.append(str(item) + '_merged')
                except AttributeError:
                    new_list.append(item)
        return new_list

    list_invoice_numbr_with_merged = convert_and_append_merged_safe(list_invoice_numbr)


    def rename_pdf_files(directory_invoice, list_invoice_numbr, dict_invoice_name_number, dict_invoice_name_posting_date):
        pdf_rename =0
        # Iterate through all files in the directory
        for filename in os.listdir(directory_invoice):
            # Check if the file is a PDF and its name (without extension) is in list_invoice_name
            if filename.endswith(".pdf") and os.path.splitext(filename)[0] in list_invoice_numbr:
                # Get the invoice name (filename without extension)
                invoice_numbr = os.path.splitext(filename)[0]
                invoice_numbr = invoice_numbr.replace('_merged', '')
                        
                # Get the corresponding value from dict_invoice_name_number
                value = dict_invoice_name_number.get(str(invoice_numbr), "")  
                posting_date = dict_invoice_name_posting_date.get(str(invoice_numbr), "")                   
            
                new_filename = f"{value}_{invoice_numbr}_{posting_date}.pdf" # Construct the new filename
            
                # Get the full paths for the old and new filenames
                old_file_path = os.path.join(directory_invoice, filename)
                new_file_path = os.path.join(directory_invoice, new_filename)

                try:
                    os.rename(old_file_path, new_file_path) # Rename the file         
                    print(f"Renamed: {filename} -> {new_filename}")
                    pdf_rename +=1
                except Exception as e:
                    print(f"failed to rename {filename}: {e}")
        print(f'done rename {pdf_rename} pdf files')
        return pdf_rename        

    directory_invoice = r"C:/Users/"+username+"/Downloads/esker_merged/" # Directory containing the PDF files. 'Documents'

    try:
        rename_pdf_files(directory_invoice, list_invoice_numbr_with_merged, format_dict_invoice_name_number, format_dict_invoice_name_posting_date)
    except Exception as e:
        print(e)
    #print(f'done rename {len(dict_invoice_name_number)} pdf files')


    def trigger_python(filename,location,recipient,argument=None):
        """ filename - exact filename without path, format - filename.py
            location - absolute file path
            recipent - list containing emails of recipents who should be notified on error
            argument - list of arguments in the same order as the code requires or None
            Note - For mail define your mail func or comment the line number 66
        """
        from subprocess import PIPE, run
        import os
        import platform
        import datetime

        ### figuring out the whole file path
        script_path=os.path.join(location,filename)

        ### logging start time and details of srguments supplied to script
        start_time=datetime.datetime.now()
        print("\n")
        print("xx"*32,' '*32,'xx'*32)
        print("xx"*32,' '*32,'xx'*32)
        # print("xx"*81)
        print("--"*81)
        print('### Running >>',script_path+', argument='+str(argument),'<<','at',start_time)
        print("--"*81,'\n')

        ### based on OS, the trigger command may change (Only Linux, Windows, MacOS included here)
        if platform.system()=='Linux':
            if argument:
                command = ['python3',script_path]+[str(i) for i in argument]
            else:
                command = ['python3',script_path]
            result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True)
        elif platform.system()=='Windows':
            if argument:
                command = ['python',script_path]+[str(i) for i in argument]
            else:
                command = ['python',script_path]
            result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True,shell=True)
        elif platform.system()=='Darwin':
            if argument:
                command = ['python3',script_path]+[str(i) for i in argument]
            else:
                command = ['python3',script_path]
            result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True)

        ### get output and error seperately
        out=result.stdout
        error=result.stderr
        print(out)
        if error:
            print(error)

        ### calculating run time after script has ended
        run_time=round((datetime.datetime.now()-start_time).seconds/60,2)

        import requests
        ### checking machine ip on which the script was triggered (this is for your reference if you are using multiple machines to run scipt)
        ### you can comment these 2 lines if not needed
        from requests import get
        server_ip = get('https://api.ipify.org').text

        ### if script gave an error, send mail else print completion status with runtime
        if result.stderr:
            body="Hi all,\nThe script named '{}' with arguments {} did not run properly for {} on server {} \n {} \n\nRegards,\nLavesh".format(filename,str(argument),str(datetime.date.today()),server_ip,error)
            subject="Code Error {} {} ".format(script_path,str(datetime.date.today()))
            #mail_func(body,subject,recipent)
            print(body)
            print(subject)
        else:
            a=datetime.date.today()
            b=script_path+', argument='+str(argument)
            c=start_time
            d=run_time

            print("--"*81)
            print("### runtime for >>",b,"<<","was",d,"min")
            print("--"*81)
            print("\n")
    
    trigger_python(filename='app.py', location='C:/Users/john.tan/esker-pending-payment-no-blank/Scripts/', 
                   recipient=[], argument=None)
    
    return (len(dict_invoice_name_number))


if __name__ =='__main__':
    main()


