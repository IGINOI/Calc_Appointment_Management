import pandas as pd
import pywhatkit as kit

days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']

# Extract the email list for 7 days
for i in range(7):
    # Read the file of the right year
    year = str(pd.Timestamp.today().date().year)
    contact_list = pd.read_excel('Appointment_Schedule_'+str(pd.Timestamp.today().date().year)+'.ods', engine='odf', sheet_name='Contatti', header=1)
    
    # Add a column to the file in order to have the full name of the contact
    contact_list['Full Name'] = (contact_list['Cognome'] + ' ' + contact_list['Nome']).str.strip()
    
    # Read the sheet of the right month
    match pd.Timestamp.today().date().month:
        case 1:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Gennaio', header=1)
        case 2:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Febbraio', header=1)
        case 3:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Marzo', header=1)
        case 4:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Aprile', header=1)
        case 5:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Maggio', header=1)
        case 6:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Giugno', header=1)
        case 7:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Luglio', header=1)
        case 8:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Agosto', header=1)
        case 9:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Settembre', header=1)
        case 10:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Ottobre', header=1)
        case 11:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Novembre', header=1)
        case 12:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Dicembre', header=1)
        case _:
            appointments_list = pd.read_excel('Appointment_Schedule_'+year+'.ods', engine='odf', sheet_name='Gennaio', header=1)
    
    # Extract the appointments for the day i (where i goes from 0 to 6 -> today to 7 days after)
    day_of_today = days_list[(pd.Timestamp.today().date().isoweekday() - 1 + i) % 7] + ' ' + str((pd.Timestamp.today().date() + pd.Timedelta(days=i)).day)
    print(day_of_today)
    appointments_day_i = appointments_list[appointments_list['Giorno'] == day_of_today]


    #################################
    ##### First slot 09:30-11:00 ####
    #################################
    appointments_day_i_first_slot = appointments_day_i[appointments_day_i['Ora'] == '9:30-11:00']
    names_day_i_first_slot = (
        list(appointments_day_i_first_slot['Marcus']) +
        list(appointments_day_i_first_slot['Jaqueline']) +
        list(appointments_day_i_first_slot['Antonio']) +
        list(appointments_day_i_first_slot['Giuseppe'])
    )
    
    email_list_first_slot = []

    for customer in names_day_i_first_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                email_list_first_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Second slot 11:00-12:30 #####
    ###################################
    appointments_day_i_second_slot = appointments_day_i[appointments_day_i['Ora'] == '11:00-12:30']
    names_day_i_second_slot = (
        list(appointments_day_i_second_slot['Marcus']) +
        list(appointments_day_i_second_slot['Jaqueline']) +
        list(appointments_day_i_second_slot['Antonio']) +
        list(appointments_day_i_second_slot['Giuseppe'])
    )
    
    email_list_second_slot = []

    for customer in names_day_i_second_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                email_list_second_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Third slot 14:00-15:30 ######
    ###################################
    appointments_day_i_third_slot = appointments_day_i[appointments_day_i['Ora'] == '14:00-15:30']
    names_day_i_third_slot = (
        list(appointments_day_i_third_slot['Marcus']) +
        list(appointments_day_i_third_slot['Jaqueline']) +
        list(appointments_day_i_third_slot['Antonio']) +
        list(appointments_day_i_third_slot['Giuseppe'])
    )
    
    email_list_third_slot = []

    for customer in names_day_i_third_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                email_list_third_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))


    ###################################
    ##### Fourth slot 15:30-17:00 #####
    ###################################
    appointments_day_i_fourth_slot = appointments_day_i[appointments_day_i['Ora'] == '15:30-17:00']
    names_day_i_fourth_slot = (
        list(appointments_day_i_fourth_slot['Marcus']) +
        list(appointments_day_i_fourth_slot['Jaqueline']) +
        list(appointments_day_i_fourth_slot['Antonio']) +
        list(appointments_day_i_fourth_slot['Giuseppe'])
    )
    
    email_list_fourth_slot = []

    for customer in names_day_i_fourth_slot:
        for contact in  list(contact_list['Full Name']):
            if customer == contact:
                email_list_fourth_slot.append(str(int(contact_list[contact_list['Full Name'] == contact]['Numero Cellulare'].values[0])))

    
    wait_time = 10
    wa_message = 'Ciao,\nti ricordiamo il tuo appuntamento del ' + day_of_today + ' alle ore 9:30-11:00. \n A presto! CentroFit'
    print(email_list_first_slot)
    for number in email_list_first_slot:
        kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)


    wa_message = 'Ciao, ti ricordo il tuo appuntamento di ' + day_of_today + ' alle ore 11:00-12:30'
    print(email_list_second_slot)
    for number in email_list_second_slot:
        kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)
    

    wa_message = 'Ciao, ti ricordo il tuo appuntamento di ' + day_of_today + ' alle ore 14:00-15:30'
    print(email_list_third_slot)
    for number in email_list_third_slot:
        kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)


    wa_message = 'Ciao, ti ricordo il tuo appuntamento di ' + day_of_today + ' alle ore 15:30-17:00'
    print(email_list_fourth_slot)
    for number in email_list_fourth_slot:
        kit.sendwhatmsg_instantly('+'+str(number), wa_message, 4, True)