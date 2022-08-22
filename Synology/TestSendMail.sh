#!/bin/sh

taskname="Email test task";
out=$(/bin/echo "Email test output"); status=$?;
host=$(/bin/hostname);
source /etc/synoinfo.conf;
headers=`printf "From: %s - %s <%s>\r\n" "$host" "$mailfrom" "$eventuser"`;
if [ ${#eventmail2} -gt 0 ];
then
	to=`printf "%s, %s" "$eventmail1" "$eventmail2"`;
else
	to=$eventmail1;
fi;
outcome="failed";
#if [ ${status} -eq 0 ];
#then
#outcome="been completed";
#fi;
outcome=`printf "%s on %s has %s" "$taskname" "$host" "$outcome"`;
subject=`printf "%s %s" "$eventsubjectprefix" "$outcome"`;
body=`printf "Dear user,\n\n%s.\n\nTask: %s\n\nSincerely,\nSynology DiskStation\n\n%s" "$outcome" "$taskname" "$out"`;
/usr/bin/php -r "mail('$to', '$subject', '$body', '$headers');";
/usr/syno/bin/synodsmnotify "admin" "Task Event" "${outcome}";

