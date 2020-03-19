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





def main():
    parser = argparse.ArgumentParser(description='Query View Image From Server')
    parser.add_argument('--maxage', '-m', required=False, help='max age of the image in the cache in minutes.')
    parser.add_argument('--logging-level', '-l', choices=['debug', 'info', 'error'], default='error',
                        help='desired logging level (set to error by default)')
    parser.add_argument('--Email_id', '-e', required=True,
                        help='Email id to send snapshot')
    args = parser.parse_args()

    print("\n -----------------Working on your Dashboard shapshot,please be patient and do not touch anything------------------------ \n")
    user = "***"
    pwd1 = "***"

    # Set logging level based on user input, or error by default
    logging_level = getattr(logging, args.logging_level.upper())
    logging.basicConfig(filename="C:/Users/sadiqmmo/PycharmProjects/untitled4/test.log", level=logging.INFO,format='%(asctime)s  %(levelname)s -  %(message)s', datefmt='%a, %d %b %Y %H:%M:%S')
    logging.info("Sign in Name:" + user)
    # Step 1: Sign in to server.
    site_id = "BI_EDM_Dev_QA"
    if not site_id:
        site_id = ""
    tableau_auth = TSC.TableauAuth(user, pwd1, site_id="BI_EDM_Dev_QA")
    server = TSC.Server(ser)
    # The new endpoint was introduced in Version 2.5
    server.version = "2.5"

    with server.auth.sign_in(tableau_auth):
        req_option = TSC.RequestOptions()
        req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,
                                         TSC.RequestOptions.Operator.Equals, 'Run Summary'))
        all_views, pagination_item = server.views.get(req_option)
        if not all_views:
            raise LookupError("View with the specified name was not found.")
        view_item = all_views[0]

        max_age = args.maxage
        if not max_age:
            max_age = 1

        image_req_option = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High,
                                                   maxage=max_age)
        #image_req_option.vf('Parameter 1', 'PROPHUB')

        server.views.populate_image(view_item, image_req_option)
        logging.info('**** View Name/Dashboard Name:' + 'run')

        with open(save_image, "wb") as image_file:
            image_file.write(view_item.image)

        ##print("View image saved to {0}".format(save_image))




        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = args.Email_id
        mail.Subject = 'EDM RUN SUMMARY SNAPSHOT'
        Body = 'Message body'

        attachment = mail.Attachments.Add(attachment_location)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
        body = "Dear" + "&nbsp;&nbsp;" + "Team" + ",""<br><br>" + "Please find the attached Latest Run run summary snapshot  " + "<br><br> <img src=""cid:MyId1"" height=""700"" width=""1150"">"
        mail.HTMLBody = (body)
        mail.Send()

        print("\n -----------------Mail Send------------------------ \n")




if __name__ == '__main__':
    main()


