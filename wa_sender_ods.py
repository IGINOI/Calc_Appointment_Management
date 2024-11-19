import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

###########################
## FUNCTIONS DEFINITION ###
###########################

# extract_names_per_slot -> used to extract the names of the customers for a specific time slot
def extract_names_per_slot(appointments_day_i, time_slot, employee_names):
    appointments_day_i_slot_i = appointments_day_i[appointments_day_i['Ora'] == time_slot]
    names_day_i_slot_i = []
    for employee in employee_names:
        names_day_i_slot_i += list(appointments_day_i_slot_i[employee])
    return (names_day_i_slot_i)

# extract_phone_numbers_per_slot -> used to extract the phone numbers of the customers for a specific time slot
def extract_phone_numbers_per_slot(names_day_i_slot_i, contact_list):
    phone_numbers_list_slot_i = []
    for customer in names_day_i_slot_i:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_slot_i.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))
    return phone_numbers_list_slot_i

# send_message -> used to send the messagges using whatsapp
def send_message(phone_number_list, appointment_day, appointment_time_slot):
    wa_message = 'Ciao, ti ricordiamo il tuo appuntamento del ' + appointment_day + ' alle ore ' + appointment_time_slot + '. A presto! CentroFit'
    for number in phone_number_list:
        # kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)
        print(wa_message)


#####################################
#### DAYS EXTRACTION FROM NOW ON ####
#####################################

# Create arrays used in the code to avoid hardcoding
months_list = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
employee_names = ['Marcus', 'Jaqueline', 'Antonio', 'Giuseppe']

# Actual extraction of the today and the days from now on
position_today = days_list.index(days_list[(pd.Timestamp.today().date().isoweekday()) % 7 - 1])

range_of_days_to_visualize = 14
week_days_from_now_on = []
for i in range(range_of_days_to_visualize):
    week_day = days_list[(pd.Timestamp.today().date().isoweekday()+i) % 7] 
    month_day = str((pd.Timestamp.today().date() + pd.Timedelta(days=i+1)).day)
    month = months_list[((pd.Timestamp.today().date() + pd.Timedelta(days=i+1)).month-1)%12]
    week_days_from_now_on.append(week_day + ' ' + month_day + ' ' + month)


###########################
####### GUI LOGIC #########
###########################
selected_indexes = []

def start_gui():
    def on_toggle(index):
        # Toggle the index in the selected list
        adjusted_index = index + 1  # Adjust index to start from 1
        if adjusted_index in selected_indexes:
            selected_indexes.remove(adjusted_index)
            toggle_vars[index].set(0)  # Unselect the toggle button
        else:
            selected_indexes.append(adjusted_index)
            toggle_vars[index].set(1)  # Select the toggle button

    def send_message():
        # Save the selected indexes and close the GUI
        root.destroy()

    def on_close():
        # Ask the user for confirmation before closing
        if messagebox.askyesno("Conferma di uscita", "Sicuro di voler uscire?"):
            selected_indexes.clear()  # Clear the selected indexes
            root.destroy()  # Close the window

    root = tk.Tk()
    root.title("Selezione giorni della quale mandare avviso")
    #root.geometry("500x400")
    root.protocol("WM_DELETE_WINDOW", on_close)

    # Create a style for the GUI
    style = ttk.Style()
    style.configure("TCheckbutton", font=("Helvetica", 12))
    style.configure("TButton", font=("Helvetica", 14))

    # Add a label at the top
    title_label = ttk.Label(root, text="Avvisa le persone con appuntamento nei seguenti giorni: ", font=("Helvetica", 14, "bold"))
    title_label.pack(pady=10)

    # Create a frame for the day selection grid
    grid_frame = ttk.Frame(root)
    grid_frame.pack(pady=10)

    # Create toggle buttons for each day in a grid layout
    toggle_vars = []
    for i, day in enumerate(week_days_from_now_on):
        var = tk.IntVar(value=0)  # 0 = off, 1 = on
        toggle_vars.append(var)
        toggle_button = ttk.Checkbutton(grid_frame, text=day, variable=var, command=lambda idx=i: on_toggle(idx))
        toggle_button.grid(row=i % 7, column=i // 7, padx=10, pady=5, sticky='w')  # Arrange in columns of 7

    # Add "Send Messages" button at the bottom
    send_button = ttk.Button(root, text="Invia Messaggi", command=send_message)
    send_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    start_gui()


#####################################
### COMPUTATION OF THE WANTED DAY ###
#####################################

wanted_days = []

for i in selected_indexes:
    if i==0:
        for j in range(1,range_of_days_to_visualize+1):
            wanted_days.append(((j) - position_today) % range_of_days_to_visualize + 1)
    else:
        wanted_days.append((int(i) - position_today) % range_of_days_to_visualize + 1)
wanted_days.sort()


#####################################
## ITERATION OVER THE WANTED DAYS ###
#####################################

# Loop over the wanted days to send the messages
for i in wanted_days:
    # Find: year - month - month_day - week_day - appointment_day(in the wanted format) for each of the next 7 days
    year = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).year
    month = months_list[(pd.Timestamp.today().date() + pd.Timedelta(days=i)).month -1]
    month_day = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).day
    week_day = days_list[(pd.Timestamp.today().date().isoweekday() - 1 + i) % 7]
    appointment_day = week_day + ' ' + str(month_day)
    
    # Read the contact list
    contact_list = pd.read_excel('Appointment_Schedule_' + str(year) + '.ods', engine='odf', sheet_name='Contatti', header=1)
    # Add a new column with the full name
    contact_list['Full Name'] = (contact_list['Cognome'] + ' ' + contact_list['Nome']).str.strip()

    # Read file of the right year and the sheet of the right month
    appointments_list = pd.read_excel('Appointment_Schedule_'+ str(year) +'.ods', engine='odf', sheet_name = month, header=1)

    # Extract the appointments for the day i
    appointments_day_i = appointments_list[appointments_list['Giorno'] == appointment_day]
    time_slots_list = list(appointments_day_i['Ora'])

    # Extract the phone numbers
    phone_numbers_day_i = []
    for i in range(len(time_slots_list)):
        names_day_i_slot_i = extract_names_per_slot(appointments_day_i, time_slots_list[0], employee_names)
        phone_numbers_list_slot_i = extract_phone_numbers_per_slot(names_day_i_slot_i, contact_list)
        phone_numbers_day_i.append(phone_numbers_list_slot_i)

    # Send whatsapp messages
    for i in range(len(time_slots_list)):
        send_message(phone_numbers_day_i[i], appointment_day, time_slots_list[i])