#!/bin/bash

host=""
user=""
cmd=`ssh $user@$host cat a.txt`

op=`echo "$cmd"|awk '{print $1}'`
port=`echo "$cmd"|awk '{print $2}'`


if [[ "$op" == "conn" ]]; then
    id=`ps aux |grep "qTfNn"|grep -v grep|awk '{print $2}'`
    if [[ -z $id ]]; then
        ssh -qTfNn -R "[::]:$port:localhost:22" $user@$host
    fi
elif [[ "$op" == "connless" ]]; then
    id=`ps aux |grep "qTfNn"|grep -v grep|awk '{print $2}'`
    if [[ -z $id ]]; then
        ssh -qTfNn -R $port:localhost:22 $user@$host
    fi
elif [[ "$op" == "disc" ]]; then
    id=`ps aux |grep "qTfNn"|grep -v grep|awk '{print $2}'`
    kill -KILL $id 2>/dev/null 1>/dev/null
fi
