#!/usr/bin/env python
# -*- coding: utf-8 -*-python

import os
import sys
import pathlib
import datetime
import xml.etree.ElementTree as ET

class CaseDocument:
    def __init__( self ):
        
        pathTMP = os.path.dirname(__file__)
        pathTMP = pathTMP.replace("modules","")
        pathTMP = pathTMP.replace("\\","/")
        pathTMP = pathTMP +u"templates/"
        fileName = pathTMP + u"template.csd"
        try:
            elementTree = ET.parse( fileName )  
        except Exception as e: 
            print (e)
        else:
            self.xmlData = elementTree.getroot ( )
            self.PrintXMLData( self.xmlData )
        return
   
    def __del__( self ):
        pass
        return

    def PrintXMLData ( self, xmlData ):
        for elem in xmlData:
            print ( elem.tag,' | ',elem.attrib )
            for subelem in elem:
                print ( 
                    subelem.tag,' | ',
                    subelem.attrib,' | ',
                    subelem.text )
        return
    
    def ReadCaseFile( self, fileName ):
        try:
            if self.xmlData:
                del self.xmlData
        except Exception as e: 
            print (e)        
        else:
            elementTree = ET.parse( fileName )
            self.xmlData = elementTree.getroot ( )
            self.PrintXMLData( self.xmlData )
        return

    def WriteCaseFile( self, fileName ):
        try:
            oldfile = pathlib.Path( fileName )
            if oldfile.is_file():
                oldfile.unlink()

            with open( fileName, 'wb' ) as fout:
                elementTree = ET.ElementTree( self.xmlData )
                elementTree.write(fout, encoding='utf-8', xml_declaration=True)

        except Exception as e: 
            print (e)
        else:
            fout.close()
        return


       
