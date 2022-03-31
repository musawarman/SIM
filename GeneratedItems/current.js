
var months=new Array(13);
months[1]="January";
months[2]="February";
months[3]="March";
months[4]="April";
months[5]="May";
months[6]="June";
months[7]="July";
months[8]="August";
months[9]="September";
months[10]="October";
months[11]="November";
months[12]="December";
var time=new Date();
var lmonth=months[time.getMonth() + 1];
var date=time.getDate();
var year=time.getYear();
if (year<100) year="20" + time.getYear();
else year=time.getYear();
document.write("<font color=#8B4513 center>" + lmonth + " ");
document.write(date + ", " + year + "</center>");

