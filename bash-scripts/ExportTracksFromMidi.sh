#!/bin/bash
# lists the first instrument used after a new track is started
if [[ $# -lt 1 ]] 
then
    echo "Usage: $0 midifile"
    exit -1
fi

strings -a -tx $1 | grep MTrk -A1 | grep -v MTrk | sed -e '/^--/d' | awk '{print substr($0, index($0, $2))}'
