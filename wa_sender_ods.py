import pandas as pd
import pywhatkit as kit

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
######## MAIN PART ########
###########################

# Create arrays used in the code to avoid hardcoding
months_list = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
#time_slots_list = ['8:30-10:00', '10:30-12:00', '12:00-13:30', '14:30-16:00', '16:00-17:30', '17:30-19:00']

ciao = input('1-Lunedì, 2-Martedì, 3-Mercoledì, 4-Giovedì, 5-Venerdì, 6-Sabato, 7-Domenica.\nInserisci il numero che corrisponde al giorno: ')

# Loop over the next 7 days to send the messages
for i in range(7):
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