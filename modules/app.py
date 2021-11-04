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
            if attrFilter is None:
                value = item.text
                if ( tagFilter is not None
                and item.tag != tagFilter ):
                    continue
                print( 'name: {}; text: {}'\
                    .format( item.tag, value ) )
            else:
                value = item.attrib[ attrFilter ]
                if ( tagFilter is not None
                and item.tag != tagFilter ):
                    continue
                print( 'name: {}; text: {}'\
                    .format( item.tag, value ) )
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
            attrName=None, itemID=False ):

        xmlData = self.elementTree.getroot( )
        for itSection in xmlData:
            if (itSection.attrib[ 'name' ] == section ):
                break
        for itSource in itSection:
            if (itSource.attrib[ 'name' ] == source ):
                break
        
        result = []
        for item in itSource:
            if attrName is None: data = item.text
            else: data = item.attrib[ attrName ]
            
            if ( attrName == 'alias' and data == '~'): 
                data = item.text
            
            if data is None: data = ''

            if itemID is True: 
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
        
        dataSource = DataSource( )
        caseData = caseDoc.elementTree.getroot( )
        rawxml = str( srcPath.read_text( encoding="utf8" ) )
        
        for section in caseData:
            for item in section:

                if ( item.text == None ): item.text = ''
                sName = section.attrib[ 'name' ]
                iName = item.attrib[ 'name' ]
                designation = f'{sName}&amp;&amp;{iName}'
                replacement = ''

                if ( sName == 'special' and iName == 'areaaddr' ):
                    try:
                        replacement = caseDoc.GetTagValue( 'area','addr' )
                        if ( ( replacement is None ) or ( replacement == '' ) ):
                            replacement = caseDoc.GetTagValue( 'area','place' )
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'special' and iName == 'clientname' ):
                    try:
                        clienttype = caseDoc.GetTagValue( 'common','clienttype' )
                        if ( clienttype == '2' ):
                            shorttype = caseDoc.GetTagValue( 'firm','shorttype' )
                            firmtitle = caseDoc.GetTagValue( 'firm','firmtitle' )
                            replacement = '{} "{}"'.format( shorttype, firmtitle )
                        if ( ( clienttype == '1' ) or ( clienttype == '3' ) ):
                            name1 = caseDoc.GetTagValue( 'person','name1' )
                            name2 = caseDoc.GetTagValue( 'person','name2' )
                            name3 = caseDoc.GetTagValue( 'person','name3' )
                            replacement = '{} {} {}'.format( name2, name1, name3 )
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'special' and iName == 'listkadnum' ):
                    try:
                        kn14 = caseDoc.GetTagValue( 'area','kn14' )
                        if ( kn14 == '' ): kn14 = None
                        if ( kn14 is not None ):
                          kn11 = caseDoc.GetTagValue( 'area','kn11' )
                          kn12 = caseDoc.GetTagValue( 'area','kn12' )
                          kn13 = caseDoc.GetTagValue( 'area','kn13' )
                          replacement = 'земельного участка {}:{}:{}:{}'\
                            .format( kn11, kn12, kn13, kn14 )
                            
                        if ( replacement == '' ):
                            kn24 = caseDoc.GetTagValue( 'area','kn24' )
                            if ( kn24 == '' ): kn24 = None
                            if ( kn24 is not None ):
                              kn21 = caseDoc.GetTagValue( 'area','kn21' )
                              kn22 = caseDoc.GetTagValue( 'area','kn22' )
                              kn23 = caseDoc.GetTagValue( 'area','kn23' )
                              replacement = 'ближайшего земельного участка {}:{}:{}:{}'\
                                .format( kn21, kn22, kn23, kn24 )
                                
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue
                    
                if ( sName == 'special' and iName == 'areakadnum' ):
                    try:
                        kn11 = caseDoc.GetTagValue( 'area','kn11' )
                        kn12 = caseDoc.GetTagValue( 'area','kn12' )
                        kn13 = caseDoc.GetTagValue( 'area','kn13' )
                        kn14 = caseDoc.GetTagValue( 'area','kn14' )                        
                        if ( kn13 == '' ): kn13 = None
                        if ( kn14 == '' ): kn14 = None
                        if ( ( kn14 is not None ) and ( kn13 is not None ) ):
                            replacement = 'земельный участок с кадастровым номером {}:{}:{}:{}'\
                                .format( kn11, kn12, kn13, kn14 )
                                
                        if ( ( kn14 is None ) and ( kn13 is not None ) ):
                            replacement = 'в границах кадастрового квартала {}:{}:{}'\
                                .format( kn11, kn12, kn13 )
                                
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'special' and iName == 'nearkadnum' ):
                    try:
                        kn21 = caseDoc.GetTagValue( 'area','kn21' )
                        kn22 = caseDoc.GetTagValue( 'area','kn22' )
                        kn23 = caseDoc.GetTagValue( 'area','kn23' )
                        kn24 = caseDoc.GetTagValue( 'area','kn24' )                        
                        if ( kn23 == '' ): kn23 = None
                        if ( kn24 == '' ): kn24 = None                        
                        if ( ( kn24 is not None ) and ( kn23 is not None ) ):
                            replacement = 'кадастровый номер ближайшего земельного участка {}:{}:{}:{}'\
                                .format( kn21, kn22, kn23, kn24 )
                                
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'special' and iName == 'anykadnum' ):
                    try:
                        kn11 = caseDoc.GetTagValue( 'area','kn11' )
                        kn12 = caseDoc.GetTagValue( 'area','kn12' )
                        kn13 = caseDoc.GetTagValue( 'area','kn13' )
                        kn14 = caseDoc.GetTagValue( 'area','kn14' )                        
                        if ( kn13 == '' ): kn13 = None
                        if ( kn14 == '' ): kn14 = None
                        if ( ( kn14 is not None ) and ( kn13 is not None ) ):
                            replacement = 'земельный участок с кадастровым номером {}:{}:{}:{}'\
                                .format( kn11, kn12, kn13, kn14 )
                                
                        if ( ( kn14 is None ) and ( kn13 is not None ) ):
                            replacement = 'в границах кадастрового квартала {}:{}:{}'\
                                .format( kn11, kn12, kn13 )
                                
                        kn21 = caseDoc.GetTagValue( 'area','kn21' )
                        kn22 = caseDoc.GetTagValue( 'area','kn22' )
                        kn23 = caseDoc.GetTagValue( 'area','kn23' )
                        kn24 = caseDoc.GetTagValue( 'area','kn24' )                        
                        if ( kn23 == '' ): kn23 = None
                        if ( kn24 == '' ): kn24 = None                        
                        if ( ( kn24 is not None ) and ( kn23 is not None ) ):
                            replacement = '{}, кадастровый номер ближайшего земельного участка {}:{}:{}:{}'\
                                .format( replacement, kn21, kn22, kn23, kn24 )
                                
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'special' and iName == 'checkdate' ):
                    try:
                        listDate = caseDoc.GetTagValue( 'common','regdate' ).split( '-' )
                        if ( len( listDate ) == 3 ):
                            day = int( listDate[ 0 ] )
                            month = int( listDate[ 1 ] )
                            year = int( listDate[ 2 ] )
                            regdate = datetime.date( year, month, day )
                            chkdate = regdate + datetime.timedelta(days=30)
                            replacement = chkdate.strftime( "%d-%m-%Y" )
                    except Exception as e: print (e)
                    else:
                        rawxml = rawxml.replace( designation, replacement )
                        print( f'designation: {designation}: text {replacement}')
                    continue

                if ( ( sName == 'common' and iName == 'cityname' )
                or ( sName == 'common' and iName == 'servname' )
                or ( sName == 'common' and iName == 'clienttype' )
                or ( sName == 'area' and iName == 'rights' )
                or ( sName == 'area' and iName == 'category' )
                or ( sName == 'area' and iName == 'terzone' ) ):                       
                    try: selection = int( item.text )
                    except Exception as e: print (e)
                    else:

                        try:
                            listReplacments = \
                                dataSource.RetrieveData( sName, iName, 'alias' )
                        except Exception as e: print (e)
                        else:
                            designation = f'{sName}&amp;&amp;{iName}@alias'        
                            replacement = listReplacments[ selection ]
                            if replacement is None: replacement = ''
                            rawxml = rawxml.replace( designation, replacement )
                            print( f'designation: {designation}: text {replacement}')

                        try:
                            listReplacments = \
                                dataSource.RetrieveData( sName, iName, 'dhead' )
                        except Exception as e: print (e)
                        else:
                            designation = f'{sName}&amp;&amp;{iName}@dhead'
                            replacement = listReplacments[ selection ]
                            if replacement is None: replacement = ''
                            rawxml = rawxml.replace( designation, replacement )
                            print( f'designation: {designation}: text {replacement}')

                        try:
                            listReplacments = \
                                dataSource.RetrieveData( sName, iName, 'fhead' )
                        except Exception as e: print (e)
                        else:
                            designation = f'{sName}&amp;&amp;{iName}@fhead'
                            replacement = listReplacments[ selection ]
                            if replacement is None: replacement = ''
                            rawxml = rawxml.replace( designation, replacement )
                            print( f'designation: {designation}: text {replacement}')

                        try:
                            listReplacments = \
                                dataSource.RetrieveData( sName, iName)
                        except Exception as e: print (e)
                        else:
                            designation = f'{sName}&amp;&amp;{iName}'                          
                            replacement = listReplacments[ selection ]
                            if replacement is None: replacement = ''
                            rawxml = rawxml.replace( designation, replacement )
                            print( f'designation: {designation}: text {replacement}')
                    continue

                if ( sName == 'area' and iName == 'spczone' ):
                    selectedSpcZones = caseDoc.GetTagValue( 
                        'area', 'spczone' ).split( ',' )
                    
                    replacement = ''    
                    if ( len( selectedSpcZones ) > 0 ):
                        try:
                            for idx, zone in enumerate(selectedSpcZones):
                                selectedSpcZones[ idx ] = ( int(zone) - 1 )
                        except Exception as e: print (e)
                        else:
                            listReplacments = \
                                    dataSource.RetrieveData( sName, iName)
                            for selection in selectedSpcZones:
                                zoneName = listReplacments[ selection ]
                                if ( replacement == ''): replacement = f'[*] {zoneName};\n'
                                else: replacement = f'{replacement}[*] {zoneName};\n'
                        
                    rawxml = rawxml.replace( designation, replacement )
                    print( f'designation: {designation}: text {replacement}')
                    continue

                replacement = item.text
                if ( sName == 'common' and iName == 'construction' ):
                    if item.text == 'True': replacement = 'имеется'
                    else: replacement = 'не имеется'

                if ( sName == 'firm' and iName == 'firmtitle' ):
                    if ( replacement != '' ):
                        replacement = f'"{replacement}"'
                        
                if replacement is None: replacement = ''
                rawxml = rawxml.replace( designation, replacement )
                print( f'designation: {designation}: text {replacement}')
        srcPath.write_text( rawxml, encoding="utf8" )        
        return
