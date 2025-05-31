import win32com.client
import os
import win32print
import win32ui
import win32con
import sys, time
import PIL.Image as Image
import PIL.ImageWin as ImageWin

    #python -m PyInstaller C:\Users\nhill_mliw72r\Documents\Code_Projects\Python\UtahOremMission\Email_Attachement_Downloader.py

def main():
    
    def find_all(name, path):
        result = []
        for root, dirs, files in os.walk(path):
            if name in files:
                result.append(os.path.join(root, name))
        return result
    
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    

    # Get the file path from user input
    print(f'\nThis Program is designed to help download email attachements.\n')
    print(f'Please save the email and put it into the folder of this application: ({base_path})\n')
    input('Press enter once you have saved the email...\n')
    while True:
        file_path =input('What is the file name? ')
        if '.msg' not in file_path:
            file_path += '.msg'
        path = os.path.join(base_path, file_path)
        print(f'searching for the file in {path}')

        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        #print('outlook initialized...')
        # Attempt to open the shared item (email)
        try:
            print('Searching...')
            msg = outlook.OpenSharedItem(path)
            print(f'File found: {msg}')
            fin_path = path
        except Exception as e:
            #p_path = os.path.join(base_path, file_path)
            parent_path = os.path.abspath(os.path.join(base_path, os.pardir))
            p_path = os.path.join(parent_path, file_path)
            print(f'Trying parent directory...{p_path}')
            try:
                msg = outlook.OpenSharedItem(p_path)
                fin_path = p_path
            except Exception as e:
                #TWO PATHS UP
                parent_path = os.path.abspath(os.path.join(parent_path, os.pardir))
                p_path = os.path.join(parent_path, file_path)
                print(f'Trying parent directory... {p_path}')
                try:
                    msg = outlook.OpenSharedItem(p_path)
                    fin_path = p_path
                except Exception as e:
                    os.system('cls' if os.name == 'nt' else 'clear')
                    print(f"The file couldn't be found, or the file is open in another application: {e}")
                    print('Please try again.')
                    continue
        break

    # Determine where to save attachments
    location = base_path

    attachements = []
    time.sleep(1)
    #os.system('cls' if os.name == 'nt' else 'clear')

    print(f'File found in: {fin_path}')
    print(f'There are {msg.Attachments.Count} attachments detected in this email')
    # Save attachments
    print(f'Saving attachements in: {base_path}')
    for att in msg.Attachments:
        attachements.append((str(att.FileName)))
        att.SaveASFile(os.path.join(location, str(att.FileName)))

    # Build the attachment paths
    for i in range(len(attachements)):
        attachements[i] = os.path.join(location, attachements[i])

    # Get the default printer
    printer_name = win32print.GetDefaultPrinter()
    # Function to print an image

    

    def print_image(file_path, two_sided):
        # Open the image using PIL
        img = Image.open(file_path)
        try:
            # Determine image orientation and rotate if needed
            img_width, img_height = img.size
            
            if img_width > img_height and horz_res < vert_res:  # Image is landscape, printer is portrait
                img = img.rotate(90, expand=True)
            elif img_height > img_width and horz_res > vert_res:  # Image is portrait, printer is landscape
                img = img.rotate(90, expand=True)

            # Get the adjusted dimensions after rotation
            img_width, img_height = img.size

            # Calculate the aspect ratio scaling
            scale_factor = min(horz_res / img_width, vert_res / img_height)
            draw_width = int(img_width * scale_factor)
            draw_height = int(img_height * scale_factor)


            # Start printing
            hdc.StartPage()
            
            # Draw the image on the printer device context
            dib = ImageWin.Dib(img)
            dib.draw(hdc.GetHandleOutput(), (0, 0, draw_width, draw_height))
            #dib.draw(hdc.GetHandleOutput(), (0, 0, horz_res, vert_res))

            #End printing
            hdc.EndPage()
            if two_sided == False:
                hdc.StartPage()
                hdc.EndPage
            
            
        except Exception as e:
            print(f"Failed to print: {e}")
            
    
    # Prompt for confirmation
    print(f'\nYour default printer is {printer_name}')
    

    # Confirm print and cleanup
    while True:
        x = input('Proceed? (Y/N) ')
        if x in ['Y', 'y', 'yes', 'Yes']:
            dev = True
            print('\nConnecting to Printer...\n')
            # Open the printer
            try:
                hprinter = win32print.OpenPrinter(printer_name)

                # Set printer defaults to allow modifications
                printer_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}

                # Open the printer with the specified defaults
                hprinter = win32print.OpenPrinter(printer_name, printer_defaults)
            except Exception:
                print('Error opening printer')
            try:
                # Get current printer properties
                properties = win32print.GetPrinter(hprinter, 2)
                devmode = properties['pDevMode']

                # Set single-sided printing (disable duplex)
                devmode.Duplex = win32con.DMDUP_SIMPLEX  # Single-sided printing

                # Apply the modified DEVMODE to the printer
                win32print.SetPrinter(hprinter, 2, properties, 0)
                
                
            except Exception:
                print('Error getting devmode...')
                dev = False
            try:
                # Create a printer device context (DC)
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer_name)

                # Get device resolution
                horz_res = hdc.GetDeviceCaps(win32con.HORZRES)
                vert_res = hdc.GetDeviceCaps(win32con.VERTRES)

                hdc.StartDoc(fin_path)

                for i in range(len(attachements)):
                    print(f'Printing image {i+1}')
                    print_image(attachements[i], dev)
                #End Printing
                hdc.EndDoc()
                win32print.ClosePrinter(hprinter)
                hdc.DeleteDC()
                for i in range(len(attachements)):
                    print(f'Deleting image {i+1}')
                    os.remove(attachements[i])
                return
            except Exception as e:
                print(f'Error printing the document: {e}')
        #elif x in ['PDF', 'pdf', 'Pdf']: #Print to pdf file
            
        elif x in ['n', 'N', 'No', 'no']:
            print('Cancelled')
            for i in range(len(attachements)):
                print(f'Deleting image {i+1}')
                os.remove(attachements[i])
            return
        else:
            print('Input Not Recognized. Please type "Yes" or "No".')
            continue

if __name__ == '__main__':
    main()
    input('Press Enter to Close The Program...')