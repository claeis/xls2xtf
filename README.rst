==============================================================================================
xls2xtf - Reads codelists from an MS-Excel file and converts it to an INTERLIS 2 transfer file
==============================================================================================

Features
========
- export codelists from an xls file
- setup an xls file, based on a INTERLIS model

License
=======
xls2xtf is licensed under the LGPL (Lesser GNU Public License).

Status
======
xls2xtf is in stable state.

System Requirements
===================
For the current version of xls2xtf, you will need a JRE (Java Runtime Environment) installed on your system, version 1.6 or later.
The JRE (Java Runtime Environment) can be downloaded for free from the Website <http://www.java.com/>.

Installing xls2xtf
==================
To install xls2xtf, choose a directory and extract the distribution file there. 

Running xls2xtf
===============
To export the codelists from the xls file, run xls2xtf with::

 java -jar xls2xtf.jar [options] file.xls file.xtf

To setup the xls file based on a model, run xls2xtf with::

 java -jar xls2xtf.jar [options] --initxls file.xls model.ili

