function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;  // First row of data to process
  var numRows = 110;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 12)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var count = 0;
  
  for (i in data) {
    count++;
    
    //KRISTA COMMENT THIS OUT WHEN YOU ARE READY TO RUN
    //if(i !=95) continue;
    

    var row = data[i];
    if(row[0] == "") continue;
    var emailAddress = row[6];  // First column
    //var emailAddress = "k.g.freeman22@gmail.com, kfreeman@andrew.cmu.edu";
    var subject = "CAM2017: Important Meeting Information";  
    var name = row[1] + " " + row[0];
    
   
    
    var introMessage = "Dear " + name + ", <br> <br>"
   + "CAM2017 is right around the corner! To help you prepare for the meeting we’ve put together some essential information. <br><br>";
   //+ "Please read this email carefully, complete any requested tasks promptly, and let us know if you have any additional questions by emailing cam2017@aps.org. <br><br>"
   //+"Also, if anyone has a nice camera and enjoys taking pictures, we are in need of a photographer during the meeting. Any help is appreciated! <br><br>";
    
    var conferenceVenue = "<b>Conference Venue </b> <br> "
    +"&nbsp;&nbsp;&nbsp; The conference will take place in the meeting rooms of the Hyatt Regency Washington on Capitol Hill: <br>" 
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; a) Address: 400 New Jersey Avenue, NW in Washington, D.C., 20001. <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; b) Phone number: +1 (202) 737-1234 <br>" 
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; c) Website (includes links to maps, parking information, and other useful facts): https://washingtondc.regency.hyatt.com/en/hotel/home.html <br> <br>";

    var transpoInfo = "<b>Getting to CAM2017 </b> <br>"
    +"&nbsp;&nbsp;&nbsp; Please see the following links for information about ground transportation to/from local airports: <br>" 
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Reagan National Airport (DCA): http://www.flyreagan.com/dca/parking-transportation <br>" 
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dulles Airport (IAD): http://www.flydulles.com/iad/parking-transportation <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; BWI Airport: https://www.bwiairport.com/to-from-bwi/transportation \n <br><br> "
    
    var conferenceDays = "<b>CAM2017 Conference Days </b> <br>"
    +"&nbsp;&nbsp;&nbsp; The following information gives you an idea of what to expect at CAM2017: <br> <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; a) You can find the conference program here: https://drive.google.com/open?id=0BxgbNpfUhZAjMWxReW5DUWtqWFE <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; We will also provide you a printed copy of this program at registration on Thursday morning. <br> <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; b) Please come to “Congressional A” at the Hyatt between 8am and 9am on Thursday morning to register and enjoy breakfast - <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;we will be getting started with the Welcoming Remarks at 9! <br> <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; c) Note that there will not be WIFI in the meeting rooms (but don’t worry, you will have it in your room free of charge!). <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Please plan ahead - if needed, be sure to download your presentation to your computer before arriving in the meeting rooms for your talk! <br><br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; d) In case you're hungry & curious, the following meals will be provided during the meeting: <br>"
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thursday 8/17: breakfast, lunch and light refreshments at the Welcome Reception (dinner on your own) <br> "
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Friday 8/18: breakfast, lunch and dinner<br> "
    +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Saturday 8/19: breakfast <br><br> ";
   
   
    var regFeeInfo = ""; 
    if(row[7] == "N"){
      regFeeInfo = "<b>Registration Fee </b> <br>"
      +"&nbsp;&nbsp;&nbsp; Our records show you have not yet paid the registration fee. <br>"
      +"&nbsp;&nbsp;&nbsp; Please do so ASAP by visiting the following website: https://www.aps.org/memb-sec/meeting/startpage.cfm?event_id=1229 <br><br>";
    }
    
    
    
    var NSFTravel = "";
    if(row[8] == "Y"){
      NSFTravel += "<b>Travel Reimbursement </b> <br>"
      +"&nbsp;&nbsp;&nbsp; Your travel award is made possible by a generous grant from the National Science Foundation. <br>"
      +"&nbsp;&nbsp;&nbsp; Please follow all guidelines for federal grant reimbursement, as detailed in this document: https://drive.google.com/open?id=0B7gPHJ17QEL1Y3ZndEZjQkJwNXc <br><br>";
    }
    
    
    
    var talkInfo = "";
    if(row[11] == "parallel"){
      talkInfo += "<b>Presentation Guidelines </b> <br>"
       +"&nbsp;&nbsp;&nbsp; You have been selected to give a talk during a parallel session. <br>" 
       +"&nbsp;&nbsp;&nbsp; Student oral presentations are 15 minutes long. We suggest 10-12 minutes for your talk and 3-5 minutes for questions. <br>" 
       +"&nbsp;&nbsp;&nbsp; Please orient your talk toward a general physics graduate student audience. <br><br>"
           +"&nbsp;&nbsp;&nbsp; To streamline the sessions, we encourage you to bring your presentation on a flash drive and transfer it to the provided laptop. <br>"
           +"&nbsp;&nbsp;&nbsp; However, if you prefer to use your own laptop, be sure to bring all the necessary equipment to connect to the projector. <br><br>";
    }
    else if(row[11] == "plenary"){
      talkInfo += "<b>Presentation Guidelines </b> <br>"
      +"&nbsp;&nbsp;&nbsp; Thank you, again, for agreeing to give an invited talk during one of our plenary sessions! <br>" 
      +"&nbsp;&nbsp;&nbsp; Plenary presentations are 40 minutes long. We suggest 30-35 minutes for your talk and 5-10 minutes for questions. <br>" 
      +"&nbsp;&nbsp;&nbsp; Please orient your talk toward a general physics graduate student audience. <br><br>"
                 +"&nbsp;&nbsp;&nbsp; To streamline the sessions, we encourage you to bring your presentation on a flash drive and transfer it to the provided laptop. <br>"
           +"&nbsp;&nbsp;&nbsp; However, if you prefer to use your own laptop, be sure to bring all the necessary equipment to connect to the projector. <br><br>";;
    }
    
    else if(row[11] == "poster"){
      talkInfo += "<b>Presentation Guidelines </b> <br>"
      +"&nbsp;&nbsp;&nbsp; You have been selected to present a poster during the poster session. <br>" 
      +"&nbsp;&nbsp;&nbsp; The poster session is 2 hours long and you should plan to stand by your poster for most of that time. <br>"
      +"&nbsp;&nbsp;&nbsp; You will have a 4’ x 8’ space to mount your poster. It is not possible to print posters at the hotel. <br><br>";
    }
    
    var chairInfo = "";
    if(row[9] != ""){
      chairInfo += "<b>Chair Guidelines </b> <br>"
        +"&nbsp;&nbsp;&nbsp; Thank you for agreeing to chair a session! You are scheduled to serve as the chair for " + row[9] + ". <br>"
        +"&nbsp;&nbsp;&nbsp; Please review the guidelines for chairs here: https://docs.google.com/document/d/1SoGp1vx3da4aZEDv3R65SKzdpg8-tUyfyAZQbH-1pZY/edit?usp=sharing <br><br>";
    }
    
    
    
    var roommateInfo = "";
    var roommateName = "";
    var roommateEmail = "";
    
    if(row[2] == "N/A")
      roommateInfo = "<b>Hotel Accommodations</b><br>"
      +"&nbsp;&nbsp;&nbsp; Our records show that you have not requested hotel accomodations. Please let us know if this is incorrect! <br><br>";
    else{
      //check for a roommate
      if(i!=0 && data[i-1][5] == row[5]){
        roommateName = data[i-1][1] + " " + data[i-1][0];
        roommateEmail = data[i-1][6];
      }
      //Logger.log(count);
      if(count < 110 && row[5] == data[count][5]){
        roommateName = data[count][1] + " " + data[count][0];
        roommateEmail = data[count][6];
      }

      roommateInfo = "<b>Hotel Accommodations</b><br>"
      +"&nbsp;&nbsp;&nbsp; Your room has been reserved at the Hyatt! Details below: <br>"
      +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; a) Your check in date is " + row[2]+ " (check in time is 3pm) and your check out date is " + row[3] + " (check out time is 12pm).<br>"
        +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; b) Your acknowledgement number is " + row[4] +" and your confirmation number is " + row[5] + ".<br>" 
        +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; c) You can check your reserved room type by entering your name and confirmation number at the following Hyatt website: <br>"
        +"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  https://www.hyatt.com/hyatt/reservations/reservation.jsp <br><br>";
      
      //if(roommateName != ""){
        //   roommateInfo += "&nbsp;&nbsp;&nbsp; Your roommate is " + roommateName + ". You can contact them at " + roommateEmail + ".<br><br>";
      //}
      
      roommateInfo += "";
      
    }

   
    var finalInfo = "As always, please feel free to contact us with any questions at cam2017@aps.org. <br><br>"
    + "Good luck in your presentation preparations - we are looking forward to meeting you in DC!" +
      "<br><br> Krista Freeman" + 
      "<br> On behalf of the CAM2017 Organizing Committee";

    //var txt = [regFeeInfo, conferenceVenue, roommateInfo, transpoInfo, conferenceDays, NSFTravel, talkInfo, chairInfo]
    var txt = [conferenceVenue, roommateInfo, conferenceDays, talkInfo, chairInfo]
    var options = {};
    var item = 0;
    options.htmlBody = introMessage;
    //options.htmlBody =  introMessage + regFeeInfo + conferenceVenue + roommateInfo+ transpoInfo + conferenceDays + NSFTravel + talkInfo + chairInfo + finalInfo;
    for (m in txt){
      if(txt[m] == "") continue;
      item++; 
      options.htmlBody += "\n \n \t" + item + ". " + txt[m]; 
    }
    
    options.htmlBody += finalInfo;
    
    //KRISTA add the name of your attachment that is in the same folder in your google drive. 
    options.attachment = [];
    options.replyTo = "cam2017@aps.org"
    
    
    
     //Arrange message
     //var message = introMessage + conferenceVenue + roommateInfo+ transpoInfo + conferenceDays + regFeeInfo + NSFTravel + talkInfo + chairInfo + finalInfo;
    
    
    //KRISTA TOGGLE THESE TWO LINES WHEN YOU ARE READY TO SEND
    //Logger.log(options)
   MailApp.sendEmail(emailAddress, subject, "", options);
  }
}
