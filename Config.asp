<%
CONST xHomePage = "../Index.Html"  '  Site Home page
CONST xdbPath = ""   '  Not used right now
CONST xUploadNotice = "paul.beaulieu@hanscom.af.mil"  ' Address to send the manditory upload Notice to  --  Should be the Webmaster
CONST xDaysToNotice = 12  ' Number of days befor sending out the request for feedback on an app that has been downloaded
CONST xDaysToReSendNotice = 5  ' Number of days befor sending out the ReRequest for feedback on an app that has been downloaded
CONST xAllowSubscrition = True  '  True = Allow people to subscribe to upload notice
DIM AdminLogin    '   Used to track if some one loged in as admin
CONST xWhatsHOT = True   '   True = Display top downloads on the main page....  False = Do not display
CONST xHotToShow = 2   '   Number of Downloads to display Only works if xWhatsHot = True





' Text Constatats
CONST xLngWhatsHotText = "TSgt Paul's All Time Top Downloads"

%>