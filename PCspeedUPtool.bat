\\Always close all applications prior to running this command.

@echo off
Cls
Title PC Quick Clean by stephen.aguilar@taskus.com
Echo WARNING: This file will delete files if you continue
Pause
Cd /d %temp%
Del *.*
Cd low
Del *.*
Cd low
Del *.*
::coded by steveo
Echo complete
Pause