#!/usr/bin/env python
# -*- coding: utf-8 -*-python

import pathlib
import datetime
import tempfile
import zipfile
import xml.etree.ElementTree as ET

class XMLParser( object ):
      
    def __init__( self ):
        return
      
    def Load( self, srcPath ):
        print( f'source file: {srcPath}')#for debug
        try:
            if self.elementTree:
                del( self.elementTree )
        except Exception as e: print (e)
        
        try:
            self.elementTree = ET.parse( srcPath )
        except Exception as e: print (e)
        return
    
    def Save( self, targetPath, xmlData = None ):
        print( f'target file: {targetPath}')#for debug
        try:
            if targetPath.exists( ): 
                targetPath.unlink( )
        except Exception as e: print (e)

        try:
            with open( targetPath, 'wb' ) as tagetFile:
                if ( xmlData != None ):
                    self.elementTree = ET.ElementTree( xmlData )                 
                self.elementTree.write( 
                    tagetFile, 
                    xml_declaration=True, 
                    encoding='utf-8', 
                    method="xml" )
            tagetFile.close( )    
        except Exception as e: print (e)
        return

    def Print( self, tagFilter = None, attrFilter = None ):
        try:
            if self.elementTree: pass
        except Exception as e: 
             print (e)
             return
        
        for item in self.elementTree.iter( ):
            if (attrFilter == None ):
                value = item.text
                if ( tagFilter == None ): print(  f'name: {item.tag}; text: {value}' )
                elif item.tag == tagFilter: print(  f'name: {item.tag}; text: {value}' )
            else:
                value = item.attrib[ attrFilter ]
                if ( tagFilter == None ): print(  f'name: {item.tag}; {attrFilter}: {value}' )
                elif item.tag == tagFilter: print(  f'name: {item.tag}; {attrFilter}: {value}' )
        return

class DataSource( XMLParser ):

    def __init__( self ):
        srcPath = pathlib.Path( __file__ )
        srcPath = srcPath.resolve( ).parent.parent
        srcPath = srcPath.joinpath( u'settings' )
        srcPath = srcPath.joinpath( u'datasrc.xml' )
        self.Load( srcPath )
        return
        
    def RetrieveData( 
            self, section, source, 
            itemAlias=False, itemID=False ):

        xmlData = self.elementTree.getroot( )
        for itSection in xmlData:
            if (itSection.attrib[ 'name' ] == section ):
                break
        for itSource in itSection:
            if (itSource.attrib[ 'name' ] == source ):
                break
        result = []
        for item in itSource:
            if not itemAlias: data = item.text
            else: data = item.attrib[ 'alias' ]
            if ( data == '~'): data = item.text
            if ( data == None ): 
                data = ''
            elif itemID: 
                dNum = item.attrib[ 'id' ]
                data = f'{dNum}) {data}'
            result.append( data )
        return result
   
class MenuQueries( XMLParser ):
    
    def __init__( self ):
        srcPath = pathlib.Path( __file__ )
        srcPath = srcPath.resolve( ).parent.parent
        srcPath = srcPath.joinpath( u'settings' )
        srcPath = srcPath.joinpath( u'queries.xml' )
        self.Load( srcPath )
        return

class DocumentCase( XMLParser ):
    
    def __init__( self ):
        srcPath = pathlib.Path( __file__ )
        srcPath = srcPath.resolve( ).parent.parent
        srcPath = srcPath.joinpath( u'templates' )
        srcPath = srcPath.joinpath( u'template.csd' )
        self.Load( srcPath )
        return

    def SetTagValue( self, sectionName, itemName, Value ):
        xmlData = self.elementTree.getroot()
        for section in xmlData:
            if ( section.attrib[ 'name' ] == sectionName ):
                break
        for item in section:
            if ( item.attrib[ 'name' ] == itemName ):
                break
        item.text = Value
        try: del( self.elementTree )
        except Exception as e: print (e)
        self.elementTree = ET.ElementTree( xmlData )
        return 

    def GetTagValue( self, sectionName, itemName ):
        xmlData = self.elementTree.getroot()
        for section in xmlData:
            if ( section.attrib[ 'name' ] == sectionName ):
                break
        for item in section:
            if ( item.attrib[ 'name' ] == itemName ):
                break
        if ( item.text == None ): return ''
        return item.text

class DOCXQuery( object ):

    def __init__( self, queID, tgtPath, caseDoc ):
        
        with tempfile.TemporaryDirectory( ) as tmpDir:

            tmpPath = pathlib.Path( tmpDir )
            print( f'Temporary dir was created: {tmpPath}')
            try:
                # копия query.zip
                quePath = pathlib.Path( __file__ )
                quePath = quePath.resolve( ).parent.parent
                quePath = quePath.joinpath( u'templates' )
                quePath = quePath.joinpath( f'query{queID}.docx' )
                zipPath = tmpPath.joinpath( u'query.zip' )
                zipPath.write_bytes( quePath.read_bytes( ) )
                # распаковка query.zip
                with zipfile.ZipFile( zipPath, 'r' ) as zipObj:
                        zipObj.extractall( path = tmpPath )
                # удаление query.zip
                if zipPath.exists( ): zipPath.unlink( )
            except Exception as e: 
                print (e)
                return

            # обновление document.xml
            srcPath = tmpPath.joinpath( u'word' )
            srcPath = srcPath.joinpath( u'document.xml' )
            self.Update( srcPath, caseDoc )

            #for debug
            #dbgPath = tgtPath.parent.joinpath( u'document.xml' )
            #dbgPath.write_bytes( srcPath.read_bytes( ) )
            
            # сохранение результата
            try:          
                with zipfile.ZipFile( 
                    tgtPath, 'w' ) as zipObj:                    
                        for itemPath in tmpPath.rglob( '*' ):
                            relPath = str( itemPath.as_posix( ) )
                            relPath = relPath.replace( 
                                str( tmpPath.as_posix( ) ) , '' )
                            relPath = relPath[ 1: ]
                            zipObj.write( itemPath.as_posix( ), relPath )
                            print( f'relation path: {relPath}' )#for debug
                zipObj.close( )
            except Exception as e: print (e)
        return
     
    def Update( self, srcPath, caseDoc ):
        
        rawxml = str( srcPath.read_text( encoding="utf8" ) )
        
        dataSource = DataSource()
        listCityNames = dataSource.RetrieveData(
            'common','cityname')
        listServNames = dataSource.RetrieveData(
            'common','servname')
        listClientTypes = dataSource.RetrieveData(
            'common','clienttype')
        listAreaRights = dataSource.RetrieveData(
            'area','right')
        listAreaCategories = dataSource.RetrieveData(
            'area','category')
        listAreaTerZones = dataSource.RetrieveData(
            'area','terzone', True)
        listAreaSpcZones = dataSource.RetrieveData(
            'area','spczone')

        caseData = caseDoc.elementTree.getroot( )
        for section in caseData:                      
            for item in section:
                if ( ( item.text != None ) 
                and ( item.text != '' ) ): 
                    sectionName = section.attrib[ 'name' ]
                    itemName = item.attrib[ 'name' ]
                    designation = f'{sectionName}&amp;&amp;{itemName}'
                    replacement = item.text
                    
                    if ( designation == 'common&amp;&amp;cityname'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listCityNames[ selection ]

                    if ( designation == 'common&amp;&amp;servname'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listServNames[ selection ]

                    if ( designation == 'common&amp;&amp;clienttype'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listClientTypes[ selection ]

                    if ( designation == 'common&amp;&amp;construction'):
                        replacement = 'имеется'

                    if ( designation == 'area&amp;&amp;rights'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listAreaRights[ selection ]

                    if ( designation == 'area&amp;&amp;category'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listAreaCategories[ selection ]

                    if ( designation == 'area&amp;&amp;terzone'):
                        try: selection = int( item.text )
                        except Exception as e: print (e)
                        else: replacement = listAreaTerZones[ selection ]

                    if ( designation == 'area&amp;&amp;spczone'):

                        selectedSpcZones = caseDoc.GetTagValue( 
                        'area', 'spczone' ).split( ',' )
                        if ( len( selectedSpcZones ) > 0 ):
                            selectedSpcZones = map( int, selectedSpcZones )
                            selectedSpcZones = [ ( x - 1 ) for x in selectedSpcZones ]
                        replacement = ''
                        for zoneIdx in selectedSpcZones:
                            selection = listAreaSpcZones[ zoneIdx ]
                            if ( replacement == ''): replacement = f'*{selection};\n'
                            else: replacement = f'{replacement}*{selection};\n'

                    print( f'designation: {designation}: text {replacement}')
                    rawxml = rawxml.replace( designation, replacement )
        srcPath.write_text( rawxml, encoding="utf8" )        
        return
