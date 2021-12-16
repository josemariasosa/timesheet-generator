#!/bin/bash

WEEKHOURS=$1
~/Repos/timesheet-generator/venv/bin/python ~/Repos/timesheet-generator/src/main.py $WEEKHOURS
