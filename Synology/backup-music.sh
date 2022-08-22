#! /bin/bash
set -x

NOW=$(date +"%Y%m%d%H%M")
CURRENTSCRIPT=$(basename $0 .sh)
LOGFILE="/log/$CURRENTSCRIPT-$NOW.log"

if sh -c ": >/dev/tty" >/dev/null 2>/dev/null; then
    TTY_ARG='--tty'
else
    TTY_ARG=''
fi

docker container run --rm --interactive $TTY_ARG \
--mount type=bind,source=/root/.aws,target=/root/.aws \
--mount type=bind,source=/volume1/Data,target=/data \
--mount type=bind,source=/volume1/Data/Management/Log,target=/log \
--env AWS_PROFILE=personal-s3-backup-writer-role \
rclone/rclone:latest \
--s3-provider AWS \
--s3-env-auth \
--s3-region eu-west-1 \
--s3-location-constraint eu-west-1 \
--s3-acl private \
--s3-storage-class STANDARD \
--log-file $LOGFILE \
--progress \
-vv \
sync /data/Music :s3:paton-backup-music