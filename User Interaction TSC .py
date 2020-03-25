import argparse
import getpass
import logging
import tableauserverclient as TSC
import win32com.client as win32
import xlrd

######### Paths ######
ser="http://awv16tableanp01"
excel_book="C:/Users/sadiqmmo/PycharmProjects/untitled4/Email.xlsx"
save_image="C:/Users/sadiqmmo/PycharmProjects/untitled4/temp.png"
attachment_location = "C:/Users/sadiqmmo/PycharmProjects/untitled4/temp.png"
######### Paths ######


print("\n -------------------Start---------------------- \n")         ##Username,Password
user = input("Please Enter your user name:")
pwd1 = getpass.getpass("Please Enter your Password:")

print("\n -------------------Snapshot options---------------------- \n")
snap1=["Direct Argument Passing based","User Interaction  based"]
for snap in range(len(snap1)):
    print("  {}.{}".format(snap, snap1[snap]))
f= int(input("Please Enter your Snapshot options from the above choices:"))
if f==0:
    g = input("Please Enter your Site id  : ")
    v= input("Please Enter your View name  : ")
    parser = argparse.ArgumentParser(description='Query View Image From Server')
    parser.add_argument('--maxage', '-m', required=False, help='max age of the image in the cache in minutes.')
    parser.add_argument('--logging-level', '-l', choices=['debug', 'info', 'error'], default='error',
                        help='desired logging level (set to error by default)')
    args = parser.parse_args()



    # Set logging level based on user input, or error by default
    logging_level = getattr(logging, args.logging_level.upper())
    logging.basicConfig(filename="C:/Users/sadiqmmo/PycharmProjects/untitled4/test.log", level=logging.INFO,format='%(asctime)s  %(levelname)s -  %(message)s', datefmt='%a, %d %b %Y %H:%M:%S')
    logging.info("Sign in Name:" + user)
    # Step 1: Sign in to server.
    site_id = g
    if not site_id:
        site_id = ""
    tableau_auth = TSC.TableauAuth(user, pwd1, site_id=g)
    server = TSC.Server(ser)
    # The new endpoint was introduced in Version 2.5
    server.version = "2.5"

    with server.auth.sign_in(tableau_auth):
        print("\n------###########$$$$$$$$$$$   List of users  $$$$$$$$$$$###########-------\n")  ##calling the function mail()
        print("\n You have the following users \n")  ##Fetching the data from the Excel file to show the list of users
        loc = (excel_book)
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
        temp200 = []
        for j in range(sheet.nrows):
            z = (sheet.cell_value(j, 1))
            temp200.append(z)

        temp300 = []
        for l in range(sheet.nrows):
            x = (sheet.cell_value(l, 0))
            temp300.append(x)
        for mail in range(len(temp200)):
            print(" {0:2d}.{1:21s}{2:4d}.{3:5s}".format(mail, temp200[mail], mail, temp300[mail]))
        e = int(input("Please Enter a user number to send a snapshot :"))
        print(
            "\n -----------------Working on your Dashboard shapshot,please be patient and do not touch anything------------------------ \n")  ##Taking the snapshot based on user input
        # Step 2: Query for the view that we want an image of
        req_option = TSC.RequestOptions()
        req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,
                                         TSC.RequestOptions.Operator.Equals, v))
        all_views, pagination_item = server.views.get(req_option)
        if not all_views:
            raise LookupError("View with the specified name was not found.")
        view_item = all_views[0]

        max_age = args.maxage
        if not max_age:
            max_age = 1

        image_req_option = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High,
                                                   maxage=max_age)
        ##image_req_option.vf('project', temp200[e])
        server.views.populate_image(view_item, image_req_option)
        logging.info('**** View Name/Dashboard Name:' + v)

        with open(save_image, "wb") as image_file:
            image_file.write(view_item.image)

        print("View image saved to {0}".format(save_image))

        loc = (excel_book)
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
        print("Sending Mail to following User:")  ##Fetching the data from the Excel file to send the particular user

        for k in range(sheet.nrows):
            a = (sheet.cell_value(e, 0))
            print(a)
            logging.info("**** Populated image for view send to:" + a)

            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = a
            mail.Subject = 'TABLEAU DASHBOARD SNAPSHOT'
            Body = 'Message body'

            attachment = mail.Attachments.Add(attachment_location)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
            body = "Dear" + "&nbsp;&nbsp;" + temp200[
                e] + ",""<br><br>" + "We appreciate your ongoing focus and commitment towards creating a high level of policy awareness and compliance in your role. On 10/11/2019 12:00:00 AM UTC, we have assigned “90 Day New Hire Learning Plan” which is due on 1/9/2020 12:00:00 AM UTC.<br><br> As a new hire, it is important for you to learn both about the industry we operate in and about Altisource’s legal and regulatory requirements" + "<br><br> <img src=""cid:MyId1"" height=""1350"" width=""1150"">"
            mail.HTMLBody = (body)
            mail.Send()

            print("\n -----------------Mail Send------------------------ \n")
            print("\n ---------------------End-------------------------- \n")
            break







elif f==1:

        parser = argparse.ArgumentParser(description='List out the names')
        parser.add_argument('--logging-level', '-l', choices=['debug', 'info', 'error'], default='error',
                            help='desired logging level (set to error by default)')
        parser.add_argument('--maxage', '-m', required=False, help='max age of the image in the cache in minutes.')
        args = parser.parse_args()
        # Set logging level based on user input, or error by default
        logging_level = getattr(logging, args.logging_level.upper())
        logging.basicConfig(filename="C:/Users/sadiqmmo/PycharmProjects/untitled4/test.log",level=logging.INFO,format='%(asctime)s  %(levelname)s -  %(message)s',datefmt='%a, %d %b %Y %H:%M:%S')

        logging.info("Sign in Name:"+user)
        tableau_auth = TSC.TableauAuth(user, pwd1)
        server = TSC.Server(ser, use_server_version=True)

        with server.auth.sign_in(tableau_auth):
            print("\n -------------------Sites---------------------- \n")      ##site listing
            print("Your server contains the following sites and Content url:\n")
            sitename = []
            site_url=[]
            for site in TSC.Pager(server.sites.get):
                sitename.append(site.name)
                site_url.append(site.content_url)

                                                                                ## Site Content Url listing
            name="Site name:-"
            content="Content Url:-"
            print("{0:>15s}{1:>42s}".format(name,content))
            for i in range(len(sitename)):
                print(" {0:2d}.{1:35s}{2:4d}.{3:5s}".format(i, sitename[i],i,site_url[i]))


            a = int(input("Please Enter a Site Number from the above choices: "))
            print("\n --------------------Projects--------------------- \n")
            tableau_auth = TSC.TableauAuth(username=user, password=pwd1, site=site_url[a])
            server = TSC.Server(ser, use_server_version=True)
            with server.auth.sign_in(tableau_auth):

                    print("Your site Number contains the following Projects:")      ## Project listing
                    temp1 = []
                    for proj in TSC.Pager(server.projects.get):
                        temp1.append(proj.name)
                    for a in range(len(temp1)):
                        print("  {}.{}".format(a, temp1[a]))


                    p1 = int(input("Enter your project Number from the above choices:"))
                    print("\n -----------------Workbooks------------------------ \n")        ## Workbooks Listing
                    print("Your Project Number contains the following Workbooks:")
                    temp10 = []
                    temp100 = []
                    for wb in TSC.Pager(server.workbooks.get):
                        temp10.append(wb.name)
                        temp100.append(wb.id)
                    for ab in range(len(temp10)):
                        print("  {}.{}".format(ab, temp10[ab]))

                                                                                            ## View Listing
                    c = int(input("Enter your workbook Number from the above choices:"))
                    print("\n -----------------Dashboards------------------------ \n")
                    print("Your workbook contains the following Dashboards:")
                    workbook = server.workbooks.get_by_id(temp100[c])
                    server.workbooks.populate_views(workbook)
                    c0 = []
                    for views in workbook.views:
                        c0.append(views.name)
                    for abc in range(len(c0)):
                        print("  {}.{}".format(abc, c0[abc]))
                    print("Your workbook contains:",len(c0),"dashboards")
                    d = int(input("Enter your Dashboard Number to take a shapshot:"))
                    print("\n -----------------Snapshot------------------------ \n")

                    print("\n------###########$$$$$$$$$$$   List of users  $$$$$$$$$$$###########-------\n")  ##calling the function mail()
                    print("\n You have the following users \n")  ##Fetching the data from the Excel file to show the list of users
                    loc = (excel_book)
                    wb = xlrd.open_workbook(loc)
                    sheet = wb.sheet_by_index(0)
                    sheet.cell_value(0, 0)
                    temp200 = []
                    for j in range(sheet.nrows):
                        z = (sheet.cell_value(j, 1))
                        temp200.append(z)

                    temp300 = []
                    for l in range(sheet.nrows):
                        x = (sheet.cell_value(l, 0))
                        temp300.append(x)
                    for mail in range(len(temp200)):
                        print(" {0:2d}.{1:21s}{2:4d}.{3:5s}".format(mail, temp200[mail], mail, temp300[mail]))
                    e = int(input("Please Enter a user number to send a snapshot :"))
                    print(
                        "\n -----------------Working on your Dashboard shapshot,please be patient and do not touch anything------------------------ \n")  ##Taking the snapshot based on user input

                    req_option = TSC.RequestOptions()
                    req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,
                                                     TSC.RequestOptions.Operator.Equals, c0[d]))
                    all_views, pagination_item = server.views.get(req_option)
                    if not all_views:
                        raise LookupError("View with the specified name was not found.")
                    view_item = all_views[0]

                    max_age = args.maxage
                    if not max_age:
                        max_age = 1

                    image_req_option = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High,
                                                               maxage=max_age)
                    ##image_req_option.vf('user param', temp200[e])
                    server.views.populate_image(view_item, image_req_option)
                    logging.info('**** View Name/Dashboard Name:' + c0[d])
                    with open(save_image, "wb") as image_file:
                        image_file.write(view_item.image)

                    print("View image saved to {0}".format(save_image))
                    loc = (excel_book)
                    wb = xlrd.open_workbook(loc)
                    sheet = wb.sheet_by_index(0)
                    sheet.cell_value(0, 0)
                    print(
                        "Sending Mail to following User:")  ##Fetching the data from the Excel file to send the particular user

                    for k in range(sheet.nrows):
                        a = (sheet.cell_value(e, 0))
                        print(a)
                        logging.info("**** Populated image for view send to:" + a)

                        outlook = win32.Dispatch('outlook.application')
                        mail = outlook.CreateItem(0)
                        mail.To = a
                        mail.Subject = 'TABLEAU DASHBOARD SNAPSHOT'
                        Body = 'Message body'

                        attachment = mail.Attachments.Add(attachment_location)
                        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
                        body = "Dear" + "&nbsp;&nbsp;" + temp200[
                            e] + ",""<br><br>" + "We appreciate your ongoing focus and commitment towards creating a high level of policy awareness and compliance in your role. On 10/11/2019 12:00:00 AM UTC, we have assigned “90 Day New Hire Learning Plan” which is due on 1/9/2020 12:00:00 AM UTC.<br><br> As a new hire, it is important for you to learn both about the industry we operate in and about Altisource’s legal and regulatory requirements" + "<br><br> <img src=""cid:MyId1"" height=""700"" width=""1150"">"
                        mail.HTMLBody = (body)
                        mail.Send()

                        print("\n -----------------Mail Send------------------------ \n")
                        print("\n ---------------------End-------------------------- \n")
                        break
else:
    print("please Enter correct choice")
