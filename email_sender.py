import pandas as pd

# Read the file for contact list and for appointment list
contact_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Contatti', header=1)
appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Gennaio', header=1)
february_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Febbraio', header=1)
march_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Marzo', header=1)
april_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Aprile', header=1)
may_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Maggio', header=1)
june_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Giugno', header=1)
july_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Luglio', header=1)
august_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Agosto', header=1)
september_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Settembre', header=1)
october_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Ottobre', header=1)
november_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Novembre', header=1)
december_appointments_list = pd.read_excel('Centro_Fit_Calendar_2025.ods', engine='odf', sheet_name='Dicembre', header=1)

#Add a column with the full name of each contact
contact_list['Full Name'] = (contact_list['Cognome'] + ' ' + contact_list['Nome']).str.strip()

days_list = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']

# Extract the email list for 7 days
for i in range(7):
    # Create the format of the day
    formatted_day = f"{days_list[i]} {i+6}"    ##!! FOR NOW THE DAYS ARE FIXED (the calendar start in january we cannot read the today date)
    print("The day " + formatted_day + " the email list is:")
    
    # Extract the appointments for the day i
    appointments_day_i = appointments_list[appointments_list['Giorno'] == formatted_day]

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
                email_list_first_slot.append(contact_list[contact_list['Full Name'] == contact]['E-mail'].values[0])


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
                email_list_second_slot.append(contact_list[contact_list['Full Name'] == contact]['E-mail'].values[0])


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
                email_list_third_slot.append(contact_list[contact_list['Full Name'] == contact]['E-mail'].values[0])


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
                email_list_fourth_slot.append(contact_list[contact_list['Full Name'] == contact]['E-mail'].values[0])

    # Print the resulting email list
    print(email_list_first_slot)
    print(email_list_second_slot)
    print(email_list_third_slot)
    print(email_list_fourth_slot)