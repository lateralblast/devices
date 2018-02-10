![alt tag](https://raw.githubusercontent.com/lateralblast/devices/master/devices.jpg)

DEVICES
=======

Diagram Export in Visio from CSV and Excel and other Sources

License
-------

This software is licensed as CC-BA (Creative Commons By Attrbution)

http://creativecommons.org/licenses/by/4.0/legalcode

Introduction
------------

A Powershell script for creating Visio Diagrams of DC racks and the hardware in them.

The import can be in the form of CSV or Excel (still under development).

This could also be used to automate Visio digram creations using exports from CMDBs (e.g. ServiceNow, Remedy, etc).

At the moment the script is a proof of concept. It has support for a number of vendor stencils and provides a framework to expand on.

Things to do:

- Excel import support

Background
----------

Originally I used the inbuilt OS applictation automation of Visio, then I tried VisioPS/Visio module.
The inbuilt OS support would not let me set the active sheet so that I could do a rack per sheet in Visio.
The VisioPS/Visio powershell module would let me set the active page correctly, but the current version 
does not appear to have the Stencil cmdlets or they have been moved in another Cmdlet and are not documented.

Thus I started using VisioBot3000 which allows me to set the active page and use Stencils:

https://github.com/MikeShepard/VisioBot3000

I have rewritten the script to utilise this powershell module.

Output
------

Example output (JPG of Visio Document) with visible stencil labels (-showlabels) and long rack names (-longrackname):

![alt tag](https://raw.githubusercontent.com/lateralblast/devices/master/rack.jpg)

Requirements
------------

The following software is required:

- Windows OS
- Powershell
- Visio
- Visio Stencils for vendor products
- VisioBot3000 Powershell Module

Installing Powershell Module:

```
Y:\Code\devices>powershell "Install-Module VisioBot3000"
```

If you've got an existing Visio Powershell Module installed, you may need to uninstall it or use the -Clobber flag to overwrite conflicting Cmdlets

If you want to clone the script and/or stencils:

- Git for Windows

Documentation
-------------

You can copy the script manually from the git repository or clone it:

```
$ git clone https://github.com/lateralblast/devices.git .
```

Stencils are put in the 'stencils' subdirectory under a 'vendor' subdirectory.

To help, I'm building a collection of Visio stencils here:

https://github.com/lateralblast/vss

This repository is getting rather large so I'd recommend you just copy the ones you need.

If you wanted to clone the entire collection:

```
$ cd devices
$ git clone https://github.com/lateralblast/vss.git stencils
```

Currently there is some support for the following vendor stencils:

- Oracle
- Dell
- Pure

Support for other vendors is relatively straight forward to add, 
you need to inspect the Visio file and look at the naming standard
for front and rear views. Common naming is "Model Front" and "Model Rear".

I plan to add some code to list the stencil names and do a match to make this process easier.

Usage
-----

To run the script from the command line you may need to alter the execution policy,
by setting it globally or adding the following command line option:

```
-ExecutionPolicy ByPass
```

Getting help:

```
Y:\Code\devices>powershell -ExecutionPolicy ByPass -File devices.ps1 -help
usage: devices.ps1
--help
--version
--inputfile  FILENAME
--outputfile FILENAME
--longracknames
--showlabels
```

Example of a CSV file:

```
$ more example.csv
Hostname,Component,Vendor,Architecture,Model,Operating System,Rack,Rack Units,Top Rack Unit,Serial Number,Asset Number,Installed Date,Warranty Exp,Location,Country
server1,Chassis,Oracle,SPARC,M3000,,A1,2,2,12341,,,,,
server2,Chassis,Oracle,SPARC,M5000,,A1,10,12,12342,,,,,
server3,Chassis,Oracle,x86,X2-4,,A1,3,15,12343,,,,,
array1,SH3,Pure,,Disk shelf,,A1,2,17,12344,,,,,
array1,SH2,Pure,,Disk shelf,,A1,2,19,12345,,,,,
array1,SH1,Pure,,Disk shelf,,A1,2,21,12346,,,,,
array1,SH0,Pure,,Disk shelf,,A1,2,23,12347,,,,,
array1,CH0,Pure,,FA-m70r2,,A1,3,26,12348,,,,,
server5,Chassis,Dell,x86,R820,,A1,2,28,12349,,,,,
flashblade1,CH1,Pure,,FlashBlade,,A1,4,32,123450
server11,Chassis,Oracle,SPARC,M3000,,A2,2,2,12351,,,,,
server12,Chassis,Oracle,SPARC,M5000,,A2,10,12,12352,,,,,
server13,Chassis,Oracle,x86,X2-4,,A2,3,15,12353,,,,,
array2,SH3,Pure,,Disk shelf,,A2,2,17,12354,,,,,
array2,SH2,Pure,,Disk shelf,,A2,2,19,12355,,,,,
array2,SH1,Pure,,Disk shelf,,A2,2,21,12356,,,,,
array2,SH0,Pure,,Disk shelf,,A2,2,23,12357,,,,,
array2,CH0,Pure,,FA-m70r2,,A2,3,26,12358,,,,,
flashblade2,CH1,Pure,,FlashBlade,,A2,4,30,123459,,,,,
```

Importing CSV file and creating Visio diagrams:

```
Y:\Code\devices>powershell -ExecutionPolicy ByPass -File devices.ps1 -inputfile input\example.csv -outputfile output\example.vsd
```
