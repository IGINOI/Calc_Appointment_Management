import pandas as pd
import pywhatkit as kit

#FUNCTION DEFINITION -> this function is used to send messagges with the pywhatkit library
def send_message(phone_number_list: list, appointment_day: str, appointment_time_slot: str):
    wa_message = 'Ciao, ti ricordiamo il tuo appuntamento del ' + appointment_day + ' alle ore ' + appointment_time_slot + '. A presto! CentroFit'
    for number in phone_number_list:
        kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)




months_list = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
time_slots_list = ['8:30-10:00', '10:30-12:00', '12:00-13:30', '14:30-16:00', '16:00-17:30', '17:30-19:00']

for i in range(7):
    # Find: year - month - month_day - week_day - appointment_day(in the wanted format)
    year = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).year
    month = months_list[(pd.Timestamp.today().date() + pd.Timedelta(days=i)).month -1]
    month_day = (pd.Timestamp.today().date() + pd.Timedelta(days=i)).day
    week_day = days_list[(pd.Timestamp.today().date().isoweekday() - 1 + i) % 7]
    appointment_day = week_day + ' ' + str(month_day)
    print(appointment_day)
    
    #Reading the contact list and adding a new column with the full name
    contact_list = pd.read_excel('Appointment_Schedule_' + str(year) + '.ods', engine='odf', sheet_name='Contatti', header=1)
    contact_list['Full Name'] = (contact_list['Cognome'] + ' ' + contact_list['Nome']).str.strip()

    #Reading the sheet of the right year and month
    appointments_list = pd.read_excel('Appointment_Schedule_'+ str(year) +'.ods', engine='odf', sheet_name = month, header=1)

    
    # Extract the appointments for the day i (where i goes from 0 to 6 -> today to 7 days after)
    
    
    appointments_day_i = appointments_list[appointments_list['Giorno'] == appointment_day]


    #################################
    ##### First slot 08:30-10:00 ####
    #################################
    appointments_day_i_first_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[0]]
    names_day_i_first_slot = (
        list(appointments_day_i_first_slot['Marcus']) +
        list(appointments_day_i_first_slot['Jaqueline']) +
        list(appointments_day_i_first_slot['Antonio']) +
        list(appointments_day_i_first_slot['Giuseppe'])
    )
    
    phone_numbers_list_first_slot = []

    for customer in names_day_i_first_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_first_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Second slot 10:30-12:00 #####
    ###################################
    appointments_day_i_second_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[1]]
    names_day_i_second_slot = (
        list(appointments_day_i_second_slot['Marcus']) +
        list(appointments_day_i_second_slot['Jaqueline']) +
        list(appointments_day_i_second_slot['Antonio']) +
        list(appointments_day_i_second_slot['Giuseppe'])
    )
    
    phone_numbers_list_second_slot = []

    for customer in names_day_i_second_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_second_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Third slot 12:00-13:30 ######
    ###################################
    appointments_day_i_third_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[2]]
    names_day_i_third_slot = (
        list(appointments_day_i_third_slot['Marcus']) +
        list(appointments_day_i_third_slot['Jaqueline']) +
        list(appointments_day_i_third_slot['Antonio']) +
        list(appointments_day_i_third_slot['Giuseppe'])
    )
    
    phone_numbers_list_third_slot = []

    for customer in names_day_i_third_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_third_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Fourth slot 14:30-16:00 #####
    ###################################
    appointments_day_i_fourth_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[3]]
    names_day_i_fourth_slot = (
        list(appointments_day_i_fourth_slot['Marcus']) +
        list(appointments_day_i_fourth_slot['Jaqueline']) +
        list(appointments_day_i_fourth_slot['Antonio']) +
        list(appointments_day_i_fourth_slot['Giuseppe'])
    )
    
    phone_numers_list_fourth_slot = []

    for customer in names_day_i_fourth_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numers_list_fourth_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Fifth slot 16:00-17:30 ######
    ###################################
    appointments_day_i_fifth_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[4]]
    names_day_i_fifth_slot = (
        list(appointments_day_i_fifth_slot['Marcus']) +
        list(appointments_day_i_fifth_slot['Jaqueline']) +
        list(appointments_day_i_fifth_slot['Antonio']) +
        list(appointments_day_i_fifth_slot['Giuseppe'])
    )
    
    phone_numbers_list_fifth_slot = []

    for customer in names_day_i_fifth_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_fifth_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Sixth slot 15:30-17:00 ######
    ###################################
    appointments_day_i_sixth_slot = appointments_day_i[appointments_day_i['Ora'] == time_slots_list[5]]
    names_day_i_sixth_slot = (
        list(appointments_day_i_sixth_slot['Marcus']) +
        list(appointments_day_i_sixth_slot['Jaqueline']) +
        list(appointments_day_i_sixth_slot['Antonio']) +
        list(appointments_day_i_sixth_slot['Giuseppe'])
    )
    
    phone_numbers_list_sixth_slot = []

    for customer in names_day_i_sixth_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                phone_numbers_list_sixth_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))

    print(phone_numbers_list_sixth_slot)
    
    phone_numbers_day_i = [phone_numbers_list_first_slot, phone_numbers_list_second_slot, phone_numbers_list_third_slot, phone_numers_list_fourth_slot, phone_numbers_list_fifth_slot, phone_numbers_list_sixth_slot]

    for i in range(6):
        send_message(phone_numbers_day_i[i], appointment_day, time_slots_list[i])