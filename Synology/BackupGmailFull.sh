#!/bin/sh

NOW=$(date +"%Y%m%d%H%M")
CURRENTSCRIPT=$(basename $0 .sh)
LOGFILE="/volume1/Data/Management/Log/$CURRENTSCRIPT-$NOW.log"

CURTIME=$(date +"%r")
ERROR_COUNT=0
CURRENT_TASK=""

report() {
	echo "$1"
	echo "$1" >> $LOGFILE
}

report "------------------------------------"
report "$CURTIME: Starting $0 ..."
report "------------------------------------"
printf "\n" >> $LOGFILE

error_handler() {
        MYSELF="$0"
		ERROR_COUNT=1
		ERROR_MESSAGE="${MYSELF}: ${CURRENT_TASK} failed: for more details check $LOGFILE"
				
		report "$ERROR_MESSAGE"
		
        /usr/syno/bin/synodsmnotify "admin" "${MYSELF}" "${ERROR_MESSAGE}"
		#/usr/bin/php -r "mail('craigpaton@gmail.com', 'Failure - $MYSELF', '$ERROR_MESSAGE', 'From: craigpaton@gmail.com');" >> $LOGFILE 2>&1
		/usr/syno/sbin/synosyslogmail --mailtype=SEVERITY --severity=ERROR --content="$ERROR_MESSAGE" >> $LOGFILE 2>&1	
}

set_task() {
	CURRENT_TASK=$1
	report "$CURRENT_TASK"
}

HOME=/var/services/homes/admin

set_task "Backup craigpaton@gmail.com"
cd /var/services/homes/admin
/volume1/Data/Backup/GMail/GMVault/gmvault_env_dsm6/bin/activate
sh /volume1/Data/Backup/GMail/GMVault/gmvault_env_dsm6/bin/gmvault sync -t full -d /volume1/Data/Backup/GMail/craigpaton@gmail.com/ craigpaton@gmail.com >> $LOGFILE 2>&1
[ $? != 0 ] && error_handler

[ $ERROR_COUNT == 0 ] && report "Success"  || report "Failed"

CURTIME=$(date +"%r")
printf "\n" >> $LOGFILE
report "------------------------------------"
report "$CURTIME: $0 finished."
report "------------------------------------"

exit $ERROR_COUNT