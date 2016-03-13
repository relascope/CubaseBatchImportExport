#!/bin/bash
# if you have a bad MusicXML Export, there will be no part-name
# and no instrument name
# in case you have a midi-program set, 
# we look up the number and get the text for it


# retrieves the Midi Number
# looks the instrument up in a table
# sets part-name and instrument-name to midi instrument

# depends-on: GM-Instrument-Map.csv, xmlstarlet, cut, grep


# Absolute path to this script.
SCRIPT=$(readlink -f $0)
# Absolute path this script is in. 
SCRIPTPATH=`dirname $SCRIPT`

if [[ $# -ne 1 ]]
then
    echo "Usage $0 MusicXmlFilename ... will be overridden!"
    exit 1
fi

xmlstarlet val "$1" &>/dev/null
if [[ $? -eq 1 ]]
then
    echo "File $1 is invalid XML. We better not touch it!"
    exit 1
fi

IFS=$'\n'
LINES=($(xmlstarlet sel -t -v "//*/midi-program/text()" $1 2>&- | xargs -i grep --null "{}" "$SCRIPTPATH/GM-Instrument-Map.csv"))

WORKING_FILE=$1

for l in "${LINES[@]}"
do
    INSTR_NAME=$(echo "$l" | cut -d , -f2)
    INSTR_NR=$(echo "$l" | cut -d , -f1)

    xmlstarlet ed -u "//*/midi-program[text()='$INSTR_NR']/../../part-name" -v $INSTR_NAME $WORKING_FILE > $WORKING_FILE.tmp
    xmlstarlet ed -u "//*/midi-program[text()='$INSTR_NR']/../../score-instrument/instrument-name" -v $INSTR_NAME $WORKING_FILE.tmp > $WORKING_FILE
done

# leave trail...
xmlstarlet ed -s "//*/encoding" -t elem -n "software" -v "http://www.dojoy.at" $WORKING_FILE > $WORKING_FILE.tmp

mv $WORKING_FILE.tmp $WORKING_FILE
