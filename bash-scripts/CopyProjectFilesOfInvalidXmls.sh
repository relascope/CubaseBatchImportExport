#!/bin/bash

if [[ $# -ne 3 ]]
then
    echo "Usage: $0 SourceDir ProjectDir DestinationDir"
    echo "   Use absolut dirs!!!"
    exit -1
fi

xmlDir=$1
cprDir=$2
destDir=$3


xmlDirSedE=$(echo "$xmlDir" | sed -e 's/[]\/$*.^|[]/\\&/g')

oldPwd=$(pwd)
cd $cprDir

$oldPwd/PrintXmlInvalidFiles.sh $1 | sed 's/\.xml/\.cpr/g' | sed "s/$xmlDirSedE//g" | sed "s/'/\\\\'/g" |  xargs -n1 -i cp --parents "{}" "$destDir"
