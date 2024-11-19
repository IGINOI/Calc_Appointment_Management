import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import ttk

###########################
### FUNCTION DEFINITION ###
###########################

# extract_names_per_slot -> used to extract the names of the customers for a specific time slot
def extract_names_per_slot(appointments_day_i, time_slot):
    appointments_day_i_slot_i = appointments_day_i[appointments_day_i['Ora'] == time_slot]
    return (
        list(appointments_day_i_slot_i['Marcus']) +
        list(appointments_day_i_slot_i['Jaqueline']) +
        list(appointments_day_i_slot_i['Antonio']) +
        list(appointments_day_i_slot_i['Giuseppe'])
    )

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
        #kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)
        print(wa_message)


###########################
######## MAIN PART 1 ######
###########################

# Create arrays used in the code to avoid hardcoding
months_list = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']

range_of_days = 14
week_days_from_now_on = []
for i in range(range_of_days):
    week_day = days_list[(pd.Timestamp.today().date().isoweekday()+i) % 7] 
    week_days_from_now_on.append(week_day + str((pd.Timestamp.today().date() + pd.Timedelta(days=i+1)).day))

selected_indexes = []
###########################
####### GUI LOGIC #########
###########################
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

    root = tk.Tk()
    root.title("Select Week Days")
    root.geometry("400x600")

    # Create toggle buttons for each day
    toggle_vars = []
    for i, day in enumerate(week_days_from_now_on):
        var = tk.IntVar(value=0)  # 0 = off, 1 = on
        toggle_vars.append(var)
        toggle_button = tk.Checkbutton(root, text=day, variable=var, onvalue=1, offvalue=0, 
                                       command=lambda idx=i: on_toggle(idx), width=12, anchor='w')
        toggle_button.grid(row=i, column=0, padx=5, pady=5, sticky='w')

    # Add "Send Messages" button
    send_button = tk.Button(root, text="Send Messages", command=send_message, bg="green", fg="white", width=15)
    send_button.grid(row=len(week_days_from_now_on), column=0, pady=10)

    root.mainloop()

if __name__ == "__main__":
    start_gui()


###########################
######## MAIN PART 2 ######
###########################

wanted_days = []
position_today = days_list.index(days_list[(pd.Timestamp.today().date().isoweekday()) % 7 - 1]) #lun=0 dom=6
print(position_today)

for i in selected_indexes:
    if i==0:
        for j in range(1,range_of_days+1):
            wanted_days.append(((j) - position_today) % range_of_days + 1)
    else:
        print("You asked for: ", i, "and today day position is: ", position_today)
        wanted_days.append((int(i) - position_today) % range_of_days + 1)
wanted_days.sort()
print("ciao", wanted_days)

# Loop over the wanted days to send the messages
for i in wanted_days:
    # Find: year - month - month_day - week_day - appointment_day(in the wanted format) for each of the next 7 days
    year = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).year
    month = months_list[(pd.Timestamp.today().date() + pd.Timedelta(days=i)).month -1]
    month_day = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).day
    week_day = days_list[(pd.Timestamp.today().date().isoweekday() - 1 + i) % 7]
    appointment_day = week_day + ' ' + str(month_day)
    print(appointment_day)
    
    # Read the contact list
    contact_list = pd.read_excel('Appointment_Schedule_' + str(year) + '.ods', engine='odf', sheet_name='Contatti', header=1)
    # Add a new column with the full name
    contact_list['Full Name'] = (contact_list['Cognome'] + ' ' + contact_list['Nome']).str.strip()

    # Read file of the right year and the sheet of the right month
    appointments_list = pd.read_excel('Appointment_Schedule_'+ str(year) +'.ods', engine='odf', sheet_name = month, header=1)

    # Extract the appointments for the day i
    appointments_day_i = appointments_list[appointments_list['Giorno'] == appointment_day]
    time_slots_list = list(appointments_day_i['Ora'])
    print(time_slots_list)

    # Extract the phone numbers
    phone_numbers_day_i = []
    for i in range(len(time_slots_list)):
        names_day_i_slot_i = extract_names_per_slot(appointments_day_i, time_slots_list[0])
        phone_numbers_list_slot_i = extract_phone_numbers_per_slot(names_day_i_slot_i, contact_list)
        phone_numbers_day_i.append(phone_numbers_list_slot_i)

    # Send whatsapp messages
    for i in range(len(time_slots_list)):
        send_message(phone_numbers_day_i[i], appointment_day, time_slots_list[i])