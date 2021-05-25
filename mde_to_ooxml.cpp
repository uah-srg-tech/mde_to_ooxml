/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/* 
 * File:   mde_to_ooxml.cpp
 * Author: user
 * 
 * Created on 16 de septiembre de 2019, 13:57
 */

#include "mde_to_ooxml.h"
#include <iostream>
#include <iomanip>
#include <sstream>
#include <cstring>
#include <minwindef.h>
#include <vector>
#include <sys/types.h>
#include <sys/stat.h>
#include <cstdio>
#include <cstdlib>
#include <cmath>
#include <dirent.h>
#include <fcntl.h>   // open
#include <unistd.h>  // read, write, close
#include <zip.h>
#include <errno.h>

using namespace std;

mde_to_ooxml::mde_to_ooxml(const char * config_filename_c_str)
{    
    /* get current dir for ini config file */
    char currentDir[FILENAME_MAX];
#ifdef __CYGWIN__
    getwd(currentDir);
#else
    getcwd(currentDir, FILENAME_MAX);
#endif
    
    if(config_filename_c_str == NULL)
    {
        config_filename = (string)currentDir + "\\mde_to_ooxml.ini";
    }
    else
    {
        ostringstream config_filename_ss;
        config_filename_ss << config_filename_c_str;
        config_filename = config_filename_ss.str();
    }
    
    status = NO_ERROR;
    
    bookmarks = 0;
    relationships = 6;//needed 5 relationships for document files
    numIds = 1;
    figures = 1;
    tables = 1;
}

uint32_t mde_to_ooxml::createDOCXfromMDE(string& ddoc_filename)
{
    ifstream configFile;
    string baseDoc_location;
    configFile.open(config_filename);
    if(!configFile.is_open())
    {
        status = CANT_OPEN_CONFIG_FILE;
        return 1;
    }
    if(!getline(configFile, baseDoc_location))
    {
        status = CONFIG_FILE_NOT_VALID;
        configFile.close();
        return 1;
    }
    if(!getline(configFile, ddoc_filename))
    {
        status = CONFIG_FILE_NOT_VALID;
        configFile.close();
        return 1;
    }
    configFile.close();
    
    //check if exists baseDoc reference folder
    struct stat info;
    if(stat(baseDoc_location.c_str(), &info) != 0)
    {
        wrong_value = baseDoc_location;
        status = CANT_STAT_DIR;
        return 1;
    }
    else if((info.st_mode & S_IFDIR) == 0)
    {
        wrong_value = baseDoc_location;
        status = FOLDER_IS_NOT_A_DIR;
        return 1;
    }
        
    size_t pos = ddoc_filename.find_last_of("\\");
    size_t extPos = ddoc_filename.substr(pos+1).find_last_of(".");
    ddoc_file_location = ddoc_filename.substr(0, pos);
    ddoc_file_name = ddoc_filename.substr(pos+1, extPos);//remove extension
    extPos = ddoc_file_name.find_last_of(".");
    if((extPos != string::npos) && (ddoc_file_name.substr(extPos).compare(".doc") == 0))
        ddoc_file_name = ddoc_file_name.substr(0, extPos); //remove second extension
    ooxml_location = ddoc_file_location + string("\\") + ddoc_file_name + string("_temp");
    
    int32_t dirStatus = 0;
#if defined (_WIN32) || defined (__CYGWIN__)
    dirStatus = mkdir(ooxml_location.c_str());
#else
    dirStatus = mkdir(ooxml_location.c_str(), 0777);
#endif
    if((dirStatus == -1) && (errno != EEXIST))
    {
        wrong_value = ooxml_location;
        status = CANT_CREATE_NEW_DIR;
        return 1;
    }
    if(chdir(ooxml_location.c_str()) != 0)
    {
        wrong_value = ooxml_location;
        status = CANT_CHANGE_DIR;
        return 1;
    }
    
    //copy minimal doc reference folder
    copyFilesInDir(baseDoc_location.c_str());
    if(status != NO_ERROR)
        return 1;
    
    //create document file
    string documentFilename = ooxml_location + "\\word\\document.xml";
    ooxml_file.open(documentFilename, ios::out|ios::trunc);
    if(ooxml_file.fail())
    {
        wrong_value = documentFilename;
        status = CANT_CREATE_OOXML_FILE;
        return 1;
    }
    
    //remove last line in numbering and rels file
    removeLastLine(ooxml_location + "\\word\\numbering.xml");
    if(status != NO_ERROR)
        return 1;
    removeLastLine(ooxml_location + "\\word\\_rels\\document.xml.rels");
    if(status != NO_ERROR)
        return 1;
    
    parseMDE(ddoc_filename);
    ooxml_file.close();
    if(status != NO_ERROR)
        return 1;
    
    //add back last tags to numbering and rels file
    addLastLine(ooxml_location + "\\word\\numbering.xml", "</w:numbering>");
    if(status != NO_ERROR)
        return 1;
    addLastLine(ooxml_location + "\\word\\_rels\\document.xml.rels", "</Relationships>");
    if(status != NO_ERROR)
        return 1;

    //zip files: create zip file in upper folder
    errorSection = string("Zipping");
    if(chdir("..") != 0)
    {
        wrong_value = string("..");
        status = CANT_CHANGE_DIR;
        return 1;
    }
    string zipName = ddoc_file_name + string(".zip"); 
    if((zipfp = zipOpen64(zipName.c_str(), 0)) == NULL)
    {
        zipClose(zipfp, NULL); 
        wrong_value = zipName;
        status = ZIPPING_ERROR;
        return 1;
    }    
    //go back to ooxml folder and zip all files and directories
    string zipLocation = ddoc_file_location + string("\\") + ddoc_file_name + string("_temp");
    if(chdir(zipLocation.c_str()) != 0)
    {
        wrong_value = zipLocation;
        status = CANT_CHANGE_DIR;
        return 1;
    }
    zipFilesInDir(zipLocation.c_str(), false);
    zipClose(zipfp, NULL); 
    if(status != NO_ERROR)
        return 1;
    
    //rename (remove old version if exists) and cleanup
    if(chdir("..") != 0)
    {
        wrong_value = string("..");
        status = CANT_CHANGE_DIR;
        return 1;
    }
    zipLocation = ddoc_file_location + (string("\\") + ddoc_file_name);
    FILE * fp;
    if((fp = fopen(string(zipLocation + ".docx").c_str(), "rb")) != NULL)
    {
        fclose(fp);
        remove(string(zipLocation + ".docx").c_str());
    }
    if(rename(string(zipLocation + ".zip").c_str(),
            string(zipLocation + ".docx").c_str()) == -1)
    {
        wrong_value = zipLocation + ".zip";
        status = RENAME_ERROR;
        return 1;
    }
    removeDir(string(zipLocation + "_temp").c_str());
    if(status != NO_ERROR)
        return 1;
    return 0;
}
    
string mde_to_ooxml::displayErrorCreateDOCXfromMDE()
{
    ostringstream error_ss;
    
    if(errorSection != "")
        error_ss << "At " << errorSection << ": ";
    
    switch(status)
    {
        case CANT_OPEN_CONFIG_FILE:
            error_ss << "Can't open config file \"" << config_filename << "\"";
            break;
            
        case CONFIG_FILE_NOT_VALID:
            error_ss << "Config file \"" << config_filename << "\" is not valid";
            break;
            
        case CANT_CREATE_OOXML_FILE:
            error_ss << "Can't create OOXML file \"" << wrong_value << "\"";
            break;
            
        case CANT_OPEN_OOXML_FILE:
            error_ss << "Can't open OOXML file \"" << wrong_value << "\"";
            break;
            
        case CANT_STAT_DIR:
            error_ss << "Can't stat dir \"" << wrong_value << "\"";
            break;
            
        case FOLDER_IS_NOT_A_DIR:
            error_ss << "Selected folder is not a dir \"" << wrong_value << "\"";
            break;
            
        case CANT_CREATE_NEW_DIR:
            error_ss << "Can't create new dir \"" << wrong_value << "\"";
            break;
            
        case CANT_CHANGE_DIR:
            error_ss << "Can't change dir \"" << wrong_value << "\"";
            break;
            
        case CANT_OPENDIR_DIR:
            error_ss << "Can't opendir dir \"" << wrong_value << "\"";
            break;
            
        case FILE_NOT_FOUND:
            error_ss << "File \"" << wrong_value << "\" not found" << endl;
            break;
            
        case FILE_PARSING_ERROR:
            error_ss << "\"" << wrong_value << "\" parsing error" << endl;
            break;
            
        case WRONG_TAG:
            error_ss << "Wrong tag \"" << wrong_value << "\". Expected \"" 
                    << wrong_value_expected << "\"" << endl;
            break;
            
        case XML_PARSING_ERROR:
            XMLerrorInfo(xmlParsingStatus, auxString, MAX_AUX_STRING);
            error_ss << "XML file \"" << wrong_value << "\" parsing error: " << auxString;
            break;
            
        case ZIPPING_ERROR:
            error_ss << "Error while zipping \"" << wrong_value << "\"" << endl;
            break;
            
        case RENAME_ERROR:
            error_ss << "Can't rename \"" << wrong_value << "\"" << endl;
            break;
            
        case REMOVE_ERROR:
            error_ss << "Can't remove \"" << wrong_value << "\"" << endl;
            break;
            
        case CANT_OPEN_FILE_TO_COPY_FROM:
            error_ss << "Can't open file \"" << wrong_value << "\" to copy from" << endl;
            break;
            
        case CANT_OPEN_FILE_TO_COPY_TO:
            error_ss << "Can't open file \"" << wrong_value << "\" to copy to" << endl;
            break;
            
        case ERROR_WHILE_COPYING:
            error_ss << "Error while copying from \"" << wrong_value << "\"" << endl;
            break;
            
        default:
            error_ss << "Unknown error " << +status;
            break;
    }
    return error_ss.str();
}

void mde_to_ooxml::parseMDE(string ddoc_filename)
{
    xmlDoc = NULL;
    xmlNodePtr root = NULL, section = NULL;
    uint32_t numberOfSections = 0;
    xmlParsingStatus = 0;
    
    LIBXML_TEST_VERSION
    
    /* look for files */
    ifstream file;
    file.open(ddoc_filename.c_str());
    if(!file.is_open())
    {
        wrong_value = ddoc_filename;
        status = FILE_NOT_FOUND;
        return;
    }
    file.close();
    
    /* Open Document */
    xmlDoc = xmlParseFile(ddoc_filename.c_str());
    if (xmlDoc == NULL)
    {
        wrong_value = ddoc_filename;
        status = FILE_PARSING_ERROR;
        return;
    }
    root = xmlDocGetRootElement(xmlDoc);
    if (root == NULL)
    {
        xmlFreeDoc(xmlDoc);
        xmlCleanupParser();
        wrong_value = ddoc_filename;
        status = FILE_PARSING_ERROR;
        return;
    }
    ooxml_file << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << endl;
    ooxml_file << "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" << endl;
    ooxml_file << "\t<w:body>" << endl;
    
    GetXMLNumChildren(root, &numberOfSections);
    wrong_value = ddoc_filename;
    
    for(uint32_t currentMainSection=0; currentMainSection<numberOfSections; ++currentMainSection)
    {
        if((xmlParsingStatus = GetXMLChildElementByIndex(root,
                currentMainSection, &section)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        parseSection(section, 1, currentMainSection, numberOfSections);
        if(status != NO_ERROR)
            break;
        cout << "\r" << fixed << setprecision(2) << ((currentMainSection+1)*100.0)/numberOfSections << "%";
    }
    xmlFreeDoc(xmlDoc);
    xmlCleanupParser();
    
    if(status != NO_ERROR)
        return;
    
    ooxml_file << "\t</w:body>" << endl;
    ooxml_file << "</w:document>";
    ooxml_file.close();       
}

void mde_to_ooxml::parseSection(xmlNodePtr sectionHandle, uint32_t currentLevel,
            uint32_t parentCurrentSection, uint32_t parentSections)
{        
    if((xmlParsingStatus = GetXMLAttributeValueByName(sectionHandle, "name",
            auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    errorSection = string(auxString);
    
    ooxml_file << "\t\t<w:p><w:pPr><w:pStyle w:val=\"Heading" << currentLevel << "\"/></w:pPr>" << endl;
    ooxml_file << "\t\t\t<w:r>" << endl;
    ooxml_file << "\t\t\t\t<w:t>" << auxString << "</w:t>" << endl;
    ooxml_file << "\t\t\t</w:r>" << endl;
    ooxml_file << "\t\t</w:p>" << endl;
    
    uint32_t numberOfSectionContents = 0;
    GetXMLNumChildren(sectionHandle, &numberOfSectionContents);
    
    for(uint32_t currentSectionContent=0; currentSectionContent<numberOfSectionContents;
            ++currentSectionContent)
    {
        xmlNodePtr sectionContentHandle = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(sectionHandle,
                currentSectionContent, &sectionContentHandle)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        if((xmlParsingStatus = GetXMLAttributeValueByName(sectionContentHandle,
                "type", auxString, MAX_AUX_STRING)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        if(strncmp(&auxString[11], "DDBody", 6) == 0)
        {
            parseBody(sectionContentHandle, "\t\t");
            if(status != NO_ERROR)
                break;
        }
        else if(strncmp(&auxString[11], "DDSection", 6) == 0)
        {
            parseSection(sectionContentHandle, currentLevel+1, 0, 0);
            if(status != NO_ERROR)
                break;
        }
        if(parentCurrentSection != 0)
        {
            cout << "\r" << fixed << setprecision(2)
                    << ((parentCurrentSection*100.0)+((currentSectionContent+1)*100.0)/numberOfSectionContents)/parentSections
                    << "%";
        }
    }
}

void mde_to_ooxml::parseBody(xmlNodePtr bodyHandle, string tab)
{
    uint32_t numberOfBodyContents = 0;
    GetXMLNumChildren(bodyHandle, &numberOfBodyContents);
    
    for(uint32_t bodyContent=0; bodyContent<numberOfBodyContents; ++bodyContent)
    {
        xmlNodePtr bodyContentHandle = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(bodyHandle,
                bodyContent, &bodyContentHandle)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        if((xmlParsingStatus = GetXMLAttributeValueByName(bodyContentHandle,
                "type", auxString, MAX_AUX_STRING)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        
        if(strncmp(&auxString[14], "DParagraph", 10) == 0)
        {
            parseParagraph(bodyContentHandle, tab);
        }
        else if(strncmp(&auxString[14], "DFigureFromFile", 15) == 0)
        {
            parseFromFile(bodyContentHandle, tab, FromFileType::FIGURE);
        }
        else if(strncmp(&auxString[14], "DTableFromFile", 14) == 0)
        {
            parseFromFile(bodyContentHandle, tab, FromFileType::TABLE);
        }
        else if(strncmp(&auxString[14], "DBasicTable", 11) == 0)
        {
            parseTable(bodyContentHandle, tab);
        }
        else if(strncmp(&auxString[14], "DEnumerate", 10) == 0)
        {
            parseListContent(bodyContentHandle, tab, 0, ListContent::ENUMERATE);
        }
        else if(strncmp(&auxString[14], "DItemize", 8) == 0)
        {
            parseListContent(bodyContentHandle, tab, 0, ListContent::ITEMIZE);
        }
        if(status != NO_ERROR)
            break;
    }
}

void mde_to_ooxml::parseParagraph(xmlNodePtr paragraphHandle, string tab,
        const string &pPr)
{
    string alignment;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(paragraphHandle,
            "alignment", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        alignment = getAlignmentOOXML(auxString);
    }
    string style;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(paragraphHandle,
            "style", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        style = getStyleOOXML(auxString, false);
    }
    string indent;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(paragraphHandle,
            "indent", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        indent = getIndentOOXML(auxString, false);
    }
    ooxml_file << tab << "<w:p>" << endl;
    if((!pPr.empty()) || (!alignment.empty()) || (!style.empty()) || (!indent.empty()))
    {
	ooxml_file << tab << "\t<w:pPr>" << endl;
        if(!pPr.empty())
            ooxml_file << tab << pPr << endl;
        if(!alignment.empty())
            ooxml_file << tab << alignment << endl;
        if(!style.empty())
            ooxml_file << tab << style << endl;
        if(!indent.empty())
            ooxml_file << tab << indent << endl;
	ooxml_file << tab << "\t</w:pPr>" << endl;
    }
    bool hasName = false;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(paragraphHandle,
            "name", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        hasName = true;
        ooxml_file << "\t\t\t<w:bookmarkStart w:id=\"" << bookmarks << "\" w:name=\"" << auxString << "\"/>" << endl;
    }
    
    uint32_t numberOfParagraphContents = 0;
    GetXMLNumChildren(paragraphHandle, &numberOfParagraphContents);
    
    for(uint32_t pContent=0; pContent<numberOfParagraphContents; ++pContent)
    {
        xmlNodePtr pContentHandle = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(paragraphHandle,
                pContent, &pContentHandle)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        if((xmlParsingStatus = GetXMLAttributeValueByName(pContentHandle,
                "type", auxString, MAX_AUX_STRING)) != 0)
        {
            status = XML_PARSING_ERROR;
            return;
        }
        if(strncmp(&auxString[14], "DRun", 4) == 0)
        {
            parseRun(pContentHandle, tab);
            if(status != NO_ERROR)
                break;
        }
        else if(strncmp(&auxString[14], "DHyperlink", 10) == 0)
        {
            xmlNodePtr hyperlinkContentHandle = NULL;

            if((xmlParsingStatus = GetXMLAttributeValueByName(pContentHandle,
                    "reference", auxString, MAX_AUX_STRING)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            
            /*
            we have to convert EMF path
            //@section.4/@sectionContent.1/@sectionContent.1/@sectionContent.0/@bodyContent.1"
            into a real xpath:
            //section[4]/sectionContent[3]/sectionContent[1]/sectionContent[0]/bodyContent[1]"
             */
            string emfPath(auxString), xPath("//");
            size_t posIni = 0, posEnd = 0;
            while((posIni = emfPath.find("/@", posEnd)) != string::npos)
            {
                //append token # if no first token
                if(posEnd != 0)
                {
                    ostringstream id;
                    id << "[" << stoi(emfPath.substr(posEnd+1, posIni-posEnd-1))+1 << "]/";
                    xPath.append(id.str());
                }
                //look for token end and append token name
                posEnd = emfPath.find(".", posIni);
                xPath.append(emfPath.substr(posIni+2, posEnd-posIni-2));
            }
            //append last token #
            ostringstream id;
            id << "[" << stoi(emfPath.substr(posEnd+1, posIni-posEnd-1))+1 << "]";
            xPath.append(id.str());
            
            xmlXPathObjectPtr xPathObject = NULL;
            xmlNodePtr xPathNode = NULL;
            if((xmlParsingStatus = getXPathObject(xmlDoc, (xmlChar*)xPath.c_str(),
                    &xPathObject)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            xPathNode = xPathObject->nodesetval->nodeTab[0];
            if((xmlParsingStatus = GetXMLAttributeValueByName(xPathNode,
                    "name", auxString, MAX_AUX_STRING)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            xmlXPathFreeObject(xPathObject);
            ooxml_file << tab << "\t<w:hyperlink w:anchor=\"" << auxString << "\">" << endl;
            
            if((xmlParsingStatus = GetXMLChildElementByTag(pContentHandle,
                    "run", &hyperlinkContentHandle)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            parseRun(hyperlinkContentHandle, tab+"\t");
            ooxml_file << tab << "\t</w:hyperlink>" << endl;
            if(status != NO_ERROR)
                break;
        }
        else
        {
            wrong_value_expected = string((char*)pContentHandle->name);
            wrong_value = "run or hyperlink";
            status = WRONG_TAG;
            break;
        }
    }
    if(status != NO_ERROR)
        return;
    if(hasName)
        ooxml_file << "\t\t\t<w:bookmarkEnd w:id=\"" << bookmarks++ << "\"/>" << endl;
    ooxml_file << tab << "</w:p>" << endl;
}

void mde_to_ooxml::parseRun(xmlNodePtr runHandle, string tab)
{
    xmlNodePtr textHandle = NULL;
    uint32_t numberOfAttributes = 0, typeIsAttribute = 0;
    
    ooxml_file << tab << "\t<w:r>";
    if((xmlParsingStatus = GetXMLNumAttributes(runHandle, &numberOfAttributes)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(strncmp((char*)runHandle->name, "paragraphContent", 16) == 0)
        typeIsAttribute = 1;

    //parse run attributes
    if((numberOfAttributes-typeIsAttribute) > 0)
    {
        ooxml_file << "<w:rPr>";
        for(uint32_t attr=typeIsAttribute; attr<numberOfAttributes; ++attr)
        {
            xmlAttrPtr runAttrHandle = NULL;
            char attributeValue[7];
            if((xmlParsingStatus = GetXMLAttributeByIndex(runHandle, attr,
                    &runAttrHandle)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            if((xmlParsingStatus = GetXMLAttributeValue(runAttrHandle, 
                    attributeValue, 7)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            if(strncmp((char*)runAttrHandle->name, "color", 4) == 0)
            {
                ooxml_file << "<w:color w:val=\"" << attributeValue << "\"/>";
            }
            else if(strncmp((char*)attributeValue, "true", 4) == 0)
            {
                if(strncmp((char*)runAttrHandle->name, "bold", 4) == 0)
                {
                    ooxml_file << "<w:b/>";
                }
                else if(strncmp((char*)runAttrHandle->name, "italics", 7) == 0)
                {
                    ooxml_file << "<w:i/>";
                }
                else if(strncmp((char*)runAttrHandle->name, "underline", 9) == 0)
                {
                    ooxml_file << "<w:u w:val=\"single\"/>";
                }
            }
        }
        ooxml_file << "</w:rPr>";
        if(status != NO_ERROR)
            return;
    }
    ooxml_file << endl;
    
    //get tab if exists
    uint32_t numberOfRunChildren = 0;
    GetXMLNumChildren(runHandle, &numberOfRunChildren);
    if(numberOfRunChildren > 1)
    {
        xmlNodePtr firstChild = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(runHandle, 0, &firstChild)) != 0)
        {
            status = XML_PARSING_ERROR;
            return;
        }
        if(strncmp((char*)firstChild->name, "tab", 3) == 0)
        {
            ooxml_file << tab << "\t\t<w:tab/>" << endl;
        }
    }
    
    //get text and text content
    if((xmlParsingStatus = GetXMLChildElementByTag(runHandle, "text",
            &textHandle)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if((xmlParsingStatus = GetXMLNumAttributes(textHandle, &numberOfAttributes)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(numberOfAttributes == 0)
    {
        ooxml_file << tab << "\t\t<w:t/>";
    }
    else
    {
        if((xmlParsingStatus = GetXMLAttributeValueByName(textHandle, "content",
                auxString, MAX_AUX_STRING)) != 0)
        {
            status = XML_PARSING_ERROR;
            return;
        }
        ooxml_file << tab << "\t\t<w:t";
        if((auxString[0] == ' ') || ((auxString[strlen(auxString)-1]) == ' '))
            ooxml_file << " xml:space=\"preserve\"";
        ooxml_file << ">" << sanitize(removeQuotes(auxString)) << "</w:t>" << endl;
    }   
    ooxml_file << tab << "\t</w:r>" << endl;
}
    
void mde_to_ooxml::parseFromFile(xmlNodePtr fromFileHandle, string tab,
        FromFileType type)
{
    uint32_t width_emu = 0, height_emu = 0;
    if((xmlParsingStatus = GetXMLAttributeValueByName(fromFileHandle,
            "width", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    width_emu = stoi(removeQuotes(auxString))*EMU_PER_PIXEL;
    if((xmlParsingStatus = GetXMLAttributeValueByName(fromFileHandle,
            "height", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    height_emu = stoi(removeQuotes(auxString))*EMU_PER_PIXEL;
    if(width_emu > MAX_EMU_WIDTH)
    {
        height_emu *= double(MAX_EMU_WIDTH)/width_emu;
        width_emu = MAX_EMU_WIDTH;
    }
    if(height_emu > MAX_EMU_HEIGTH)
    {
        width_emu *= double(MAX_EMU_HEIGTH)/width_emu;
        height_emu = MAX_EMU_HEIGTH;
    }
    
    string alignment;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(fromFileHandle,
            "alignment", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        alignment = getAlignmentOOXML(auxString);
    }
    string style;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(fromFileHandle,
            "style", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        style = getStyleOOXML(auxString, false);
    }
    string indent;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(fromFileHandle,
            "indent", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        indent = getIndentOOXML(auxString, false);
    }
    
    if((xmlParsingStatus = GetXMLAttributeValueByName(fromFileHandle,
            "name", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    ooxml_file << tab << "<w:p>" << endl;
    if((!alignment.empty()) || (!style.empty()) || (!indent.empty()))
    {
	ooxml_file << tab << "\t<w:pPr>" << endl;
        if(!alignment.empty())
            ooxml_file << tab << alignment << endl;
        if(!style.empty())
            ooxml_file << tab << style << endl;
        if(!indent.empty())
            ooxml_file << tab << indent << endl;
	ooxml_file << tab << "\t</w:pPr>" << endl;
    }
    ooxml_file << tab << "\t<w:bookmarkStart w:id=\"" << bookmarks << "\" w:name=\"" << auxString << "\"/>" << endl;
    ooxml_file << tab << "\t<w:r>" << endl;
    ooxml_file << tab << "\t\t<w:drawing>" << endl;
    ooxml_file << tab << "\t\t\t<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" << endl;
    ooxml_file << tab << "\t\t\t\t<wp:extent cx=\"" << width_emu << "\" cy=\"" << height_emu << "\"/>" << endl;
    if(type == FIGURE)
        ooxml_file << tab << "\t\t\t\t<wp:docPr id=\"" << figures << "\" name=\"Figure " << figures << "\"/>" << endl;
    else if(type == TABLE)
        ooxml_file << tab << "\t\t\t\t<wp:docPr id=\"" << figures << "\" name=\"Figure " << figures << "\"/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" << endl;
    ooxml_file << tab << "\t\t\t\t\t<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t<pic:nvPicPr>" << endl;
    if((xmlParsingStatus = GetXMLAttributeValueByName(fromFileHandle,
            "referenceFile", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    
    string srcFilename, dstFilename, auxString_str;
    dstFilename = ooxml_location + string("\\word\\media\\");
    //check if media folder exists, and create it if necessary
    DIR* dir = opendir(dstFilename.c_str());
    if(dir)
    {
        closedir(dir);
    }
    else
    {
        int32_t dirStatus = 0;
#if defined (_WIN32) || defined (__CYGWIN__)
        dirStatus = mkdir(dstFilename.c_str());
#else
        dirStatus = mkdir(mediaFolder.c_str()), 0777);
#endif
        if((dirStatus == -1) && (errno != EEXIST))
        {
            wrong_value = dstFilename;
            status = CANT_CREATE_NEW_DIR;
            return;
        }
    }
    
    srcFilename = ddoc_file_location + string("\\") + auxString;
    auxString_str = string(auxString);
    size_t pos = auxString_str.find_last_of("\\");
    if(pos == string::npos)
        pos = auxString_str.find_last_of("/");
    dstFilename += auxString_str.substr(pos+1);
    
    ifstream file;
    file.open(srcFilename);
    if(!file.is_open())
    {
        wrong_value = srcFilename;
        status = FILE_NOT_FOUND;
        return;
    }
    file.close();
    
    copyFile(srcFilename, dstFilename);
    if(status != NO_ERROR)
        return;
    
    string relsFilename = ooxml_location + "\\word\\_rels\\document.xml.rels";
    ooxml_aux_file.open(relsFilename.c_str(), ios::out|ios::in|ios::ate);
    if(ooxml_file.fail())
    {
        wrong_value = relsFilename;
        status = CANT_OPEN_OOXML_FILE;
        return;
    }
    ooxml_aux_file << "\t<Relationship Id=\"rId" << relationships << "\" "
            "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
            "Target=\"media/" << auxString_str.substr(pos+1) << "\"/>" << endl;
    ooxml_aux_file.close();
            
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<pic:cNvPr id=\"1\" name=\"" << auxString << "\"/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<pic:cNvPicPr/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t</pic:nvPicPr>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t<pic:blipFill>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<a:blip r:embed=\"rId" << relationships++ << "\"/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<a:stretch><a:fillRect/></a:stretch>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t</pic:blipFill>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t<pic:spPr>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<a:xfrm>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t\t<a:off x=\"0\" y=\"0\"/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t\t<a:ext cx=\"" << width_emu << "\" cy=\"" << height_emu << "\"/>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t</a:xfrm>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t\t<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t\t</pic:spPr>" << endl;
    ooxml_file << tab << "\t\t\t\t\t\t</pic:pic>" << endl;
    ooxml_file << tab << "\t\t\t\t\t</a:graphicData>" << endl;
    ooxml_file << tab << "\t\t\t\t</a:graphic>" << endl;
    ooxml_file << tab << "\t\t\t</wp:inline>" << endl;
    ooxml_file << tab << "\t\t</w:drawing>" << endl;
    ooxml_file << tab << "\t</w:r>" << endl;
    ooxml_file << tab << "\t<w:bookmarkEnd w:id=\"" << bookmarks++ << "\"/>" << endl;
    ooxml_file << tab << "</w:p>" << endl;
    
    if((xmlParsingStatus = SearchXMLAttributeValueByName(fromFileHandle,
            "caption", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        parseCaption(auxString, type, tab, alignment);
    }
}

void mde_to_ooxml::parseTable(xmlNodePtr tableHandle, string tab)
{
    uint32_t width=0;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "width", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        width = stoi(removeQuotes(auxString))*50;
    }
    
    string alignment;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "alignment", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        alignment = getAlignmentOOXML(auxString);
    }
    string style;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "style", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        style = getStyleOOXML(auxString, true);
    }
    string indent;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "indent", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        indent = getIndentOOXML(auxString, true);
    }
    
    bool tableHasName = false;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "name", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        tableHasName = true;
    }
    ooxml_file << tab << "<w:tbl>" << endl;
    if(tableHasName)
        ooxml_file << tab << "\t<w:bookmarkStart w:id=\"" << bookmarks << "\" w:name=\"" <<  auxString << "\"/>" << endl;
    ooxml_file << tab << "\t<w:tblPr>" << endl;
    if(width > 0)
        ooxml_file << tab << "\t\t<w:tblW w:w=\"" << width << "\" w:type=\"pct\"/>" << endl;
    if(!alignment.empty())
        ooxml_file << tab << alignment << endl;
    if(style.empty())
        ooxml_file << tab << "\t\t<w:tblStyle w:val=\"TableGrid\"/>" << endl;
    else
        ooxml_file << tab << style << endl;
    if(!indent.empty())
        ooxml_file << tab << indent << endl;
    ooxml_file << tab << "\t</w:tblPr>" << endl;

    uint32_t numberOfRows = 0;
    uint32_t numberOfCellsTableWidth = 0;
    vector<int32_t> vMergeLast;
    vector<uint32_t> currentGridSpan, lastGridSpan;
    GetXMLNumChildren(tableHandle, &numberOfRows);

    for(uint32_t currentRow=0; currentRow<numberOfRows; ++currentRow)
    {
        uint32_t currentDDocCell = 0;
        xmlNodePtr rowHandle = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(tableHandle,
                currentRow, &rowHandle)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        ooxml_file << tab << "\t<w:tr>" << endl;

        if(currentRow == 0)
        {
            GetXMLNumChildren(rowHandle, &numberOfCellsTableWidth);
        }
        
        for(uint32_t currentCell=0; currentCell<numberOfCellsTableWidth; ++currentCell)
        {
            if(currentRow == 0)
            {
                /* at first row, insert vMerge info at vectors */
                vMergeLast.push_back(-1);
                currentGridSpan.push_back(1);
                lastGridSpan.push_back(1);
            }
            if((int32_t)currentRow < vMergeLast.at(currentCell))
            {
                ooxml_file << tab << "\t\t<w:tc>" << endl;
                ooxml_file << tab << "\t\t\t<w:tcPr>" << endl;
                if(currentGridSpan.at(currentCell) != 1)
                {
                    ooxml_file << tab << "\t\t\t\t<w:gridSpan w:val=\"" <<
                            currentGridSpan.at(currentCell) << "\"/>" << endl;
                }
                ooxml_file << tab << "\t\t\t\t<w:vMerge w:val=\"continue\"/>" << endl;
                ooxml_file << tab << "\t\t\t\t<w:p/>" << endl;
                ooxml_file << tab << "\t\t\t</w:tcPr>" << endl;
                ooxml_file << tab << "\t\t</w:tc>" << endl;
                continue;
            }
            if(currentGridSpan.at(currentCell) == 0)
            {
                continue;
            }
            xmlNodePtr cellHandle = NULL;
            if((xmlParsingStatus = GetXMLChildElementByIndex(rowHandle,
                    currentDDocCell, &cellHandle)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            
            uint32_t numberOfCellAttributes = 0;
            xmlAttrPtr cellAttribHandle = NULL;
            
            if((xmlParsingStatus = GetXMLNumAttributes(cellHandle,
                    &numberOfCellAttributes)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            ooxml_file << tab << "\t\t<w:tc>" << endl;
            ooxml_file << tab << "\t\t\t<w:tcPr>" << endl;
            uint32_t gridSpan = 1;
            for(uint32_t cellAttrib=0; cellAttrib<numberOfCellAttributes; ++cellAttrib)
            {
                if((xmlParsingStatus = GetXMLAttributeByIndex(cellHandle,
                        cellAttrib, &cellAttribHandle)) != 0)
                {
                    status = XML_PARSING_ERROR;
                    break;
                }
                if((xmlParsingStatus = GetXMLAttributeValue(cellAttribHandle,
                        auxString, MAX_AUX_STRING)) != 0)
                {
                    status = XML_PARSING_ERROR;
                    break;
                }
                if(!strncmp((char*)cellAttribHandle->name, "colSpan", 7))
                {
                    gridSpan = stoi(removeQuotes(auxString));
                    ooxml_file << tab << "\t\t\t\t<w:gridSpan w:val=\"" <<
                            gridSpan << "\"/>" << endl;
                }
                else if(!strncmp((char*)cellAttribHandle->name, "rowSpan", 7))
                {
                    ooxml_file << tab << "\t\t\t\t<w:vMerge w:val=\"restart\"/>" << endl;
                    vMergeLast.at(currentCell) = currentRow+stoi(removeQuotes(auxString));
                    currentGridSpan.at(currentCell) = gridSpan;
                }
                else if(!strncmp((char*)cellAttribHandle->name, "width", 5))
                {
                    ooxml_file << tab << "\t\t\t\t<w:tcW w:w=\"" << 
                            stoi(removeQuotes(auxString))*50 << "\" w:type=\"pct\"/>" << endl;
                }
                else if(!strncmp((char*)cellAttribHandle->name, "shadow", 6))
                {
                    ooxml_file << tab <<
                            "\t\t\t\t<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"" << 
                            auxString << "\"/>" << endl;
                }

                lastGridSpan.at(currentCell) = currentGridSpan.at(currentCell);
                currentGridSpan.at(currentCell) = gridSpan;
                if(currentRow == 0)
                {
                    if(gridSpan != 1)
                    {
                        numberOfCellsTableWidth += gridSpan-1;
                        for(uint32_t span=1; span<gridSpan; ++span)
                        {
                            vMergeLast.push_back(-1);
                            currentGridSpan.push_back(0);
                        }
                    }
                }
                else
                {
                   if(gridSpan != 1)
                   {
                       for(uint32_t span=1; span<gridSpan; ++span)
                       {
                           currentGridSpan.at(currentCell + span) = 0; 
                       }
                   }
                   else if(lastGridSpan.at(currentCell) != 1)
                   {
                       for(uint32_t span=1; span<lastGridSpan.at(currentCell); ++span)
                       {
                           currentGridSpan.at(currentCell + span) = 1; 
                       }
                   }
                }
                
            }
            if(status != NO_ERROR)
                break;
            ooxml_file << tab << "\t\t\t</w:tcPr>" << endl;

            parseBody(cellHandle, tab+"\t\t\t");
            if(status != NO_ERROR)
                break;
            ooxml_file << tab << "\t\t</w:tc>" << endl;
            currentDDocCell++;
        }
        ooxml_file << tab << "\t</w:tr>" << endl;
        if(status != NO_ERROR)
            break;
    }
    if(status != NO_ERROR)
        return;
    if(tableHasName)
        ooxml_file << tab << "\t<w:bookmarkEnd w:id=\"" << bookmarks++ << "\"/>" << endl;
    ooxml_file << tab << "</w:tbl>" << endl;
    
    if((xmlParsingStatus = SearchXMLAttributeValueByName(tableHandle,
            "caption", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        parseCaption(auxString, FromFileType::TABLE, tab, alignment);
    }
}

void mde_to_ooxml::parseListContent(xmlNodePtr listHandle, string tab,
        uint32_t currentListLevel, ListContent listContent)
{
    string alignment;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(listHandle,
            "alignment", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        alignment = getAlignmentOOXML(auxString);
    }
    string style;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(listHandle,
            "style", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        style = getStyleOOXML(auxString, false);
    }
    string indent;
    if((xmlParsingStatus = SearchXMLAttributeValueByName(listHandle,
            "indent", auxString, MAX_AUX_STRING)) != 0)
    {
        status = XML_PARSING_ERROR;
        return;
    }
    if(auxString[0] != '\0')
    {
        indent = getIndentOOXML(auxString, false);
    }
    
    uint32_t numberOfItems = 0;
    GetXMLNumChildren(listHandle, &numberOfItems);
    
    for(uint32_t item=0; item<numberOfItems; ++item)
    {
        xmlNodePtr itemHandle = NULL;
        if((xmlParsingStatus = GetXMLChildElementByIndex(listHandle,
                item, &itemHandle)) != 0)
        {
            status = XML_PARSING_ERROR;
            break;
        }
        uint32_t numberOfItemContents = 0;
        GetXMLNumChildren(itemHandle, &numberOfItemContents);

        for(uint32_t itemContent=0; itemContent<numberOfItemContents; ++itemContent)
        {
            xmlNodePtr itemContentHandle = NULL, itemContentChildHandle = NULL;
            if((xmlParsingStatus = GetXMLChildElementByIndex(itemHandle,
                    itemContent, &itemContentHandle)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            if((xmlParsingStatus = GetXMLChildElementByIndex(itemContentHandle,
                    0, &itemContentChildHandle)) != 0)
            {
                status = XML_PARSING_ERROR;
                break;
            }
            if(strncmp((char*)itemContentHandle->name, "paragraph", 9) == 0)
            {
                ostringstream pPr_ss;
                pPr_ss << tab << "<w:pStyle w:val=\"ListParagraph\"/><w:numPr>" <<
                        "<w:ilvl w:val=\"" << currentListLevel << "\"/>" <<
                        "<w:numId w:val=\""<< numIds << "\"/></w:numPr>";
                if(!alignment.empty())
                    pPr_ss << endl << tab << alignment;
                if(!style.empty())
                    pPr_ss << endl << tab << style;
                if(!indent.empty())
                    pPr_ss << endl << tab << indent;
                parseParagraph(itemContentHandle, tab, pPr_ss.str());
            }
            else if(strncmp((char*)itemContentHandle->name, "sublist", 7) == 0)
            {
                parseListContent(itemContentHandle, tab, currentListLevel+1,
                        listContent);
            }
            if(status != NO_ERROR)
                break;
        }
        if(status != NO_ERROR)
            break;
    }
                
    if(currentListLevel == 0)//only once per list, only if level is 0
    {
        string numberingFilename = ooxml_location + "\\word\\numbering.xml";
        ooxml_aux_file.open(numberingFilename.c_str(), ios::out|ios::in|ios::ate);
        if(ooxml_file.fail())
        {
            wrong_value = numberingFilename;
            status = CANT_OPEN_OOXML_FILE;
            return;
        }        
        ooxml_aux_file << "\t<w:num w:numId=\"" << numIds++ << "\">" << endl;
        if(listContent == ListContent::ENUMERATE)
            ooxml_aux_file << "\t\t<w:abstractNumId w:val=\"0\"/>" << endl;
        else if(listContent == ListContent::ITEMIZE)
            ooxml_aux_file << "\t\t<w:abstractNumId w:val=\"1\"/>" << endl;
        ooxml_aux_file << "\t</w:num>" << endl;
        ooxml_aux_file.close();
    }
}

void mde_to_ooxml::parseCaption(const char * captionText, FromFileType type,
        string tab, string alignment)
{
    ooxml_file << tab << "<w:p>" << endl;
    ooxml_file << tab << "\t<w:pPr>" << endl;
    ooxml_file << tab << "\t\t<w:pStyle w:val=\"Caption\"/>" << endl;
    if(!alignment.empty())
        ooxml_file << tab << alignment << endl;
    ooxml_file << tab << "\t</w:pPr>" << endl;
    ooxml_file << tab << "\t<w:r>" << endl;
    if(type == FromFileType::FIGURE)
        ooxml_file << tab << "\t\t<w:t>Figure " << figures++ << ": ";
    else if(type == FromFileType::TABLE)
        ooxml_file << tab << "\t\t<w:t>Table " << tables++ << ": ";
    ooxml_file << captionText << "</w:t>" << endl;
    ooxml_file << tab << "\t</w:r>" << endl;
    ooxml_file << tab << "</w:p>" << endl;
}

void mde_to_ooxml::removeLastLine(string filename)
{
    vector<string> textLines;
    
    //open file and read content
    ifstream in_file(filename.c_str());
    if(in_file.fail())
    {
        wrong_value = filename;
        status = CANT_OPEN_OOXML_FILE;
        return;
    }
    else
    {
        string auxLine;
        while(in_file.good())
        {
            getline(in_file, auxLine);
            textLines.push_back(auxLine);
        }
    }
    in_file.close();
    
    //remove last line
    textLines.erase(textLines.end()-1);

    //open the file again and save all vector
    ooxml_aux_file.open(filename.c_str(), ios::out|ios::trunc);
    if(ooxml_aux_file.fail())
    {
        wrong_value = filename;
        status = CANT_OPEN_OOXML_FILE;
        return;
    }
    else
    {
        for(size_t i=0; i<textLines.size(); i++)
            ooxml_aux_file << textLines.at(i) << endl;
        ooxml_aux_file.close();
    }
}

void mde_to_ooxml::addLastLine(string filename, string lineToAdd)
{
    ooxml_aux_file.open(filename.c_str(), ios::out|ios::app);
    if(ooxml_aux_file.fail())
    {
        wrong_value = filename;
        status = CANT_OPEN_OOXML_FILE;
        return;
    }
    else
    {
        ooxml_aux_file << lineToAdd;
        ooxml_aux_file.close();
    }
}

void mde_to_ooxml::copyFilesInDir(const char * dirname)
{
    DIR *dir;
    struct dirent *ent;
    if((dir = opendir(dirname)) == NULL)
    {
        wrong_value = string(dirname);
        status = CANT_OPENDIR_DIR;
        return;
    }
    else
    {
        while((ent = readdir(dir)) != NULL)
        {
            if(((strlen(ent->d_name) == 1) && (strncmp(ent->d_name, ".", 1) == 0)) ||
                    ((strlen(ent->d_name) == 2) && (strncmp(ent->d_name, "..", 2) == 0)))
            {
                continue;
            }
            string srcFilename = string(dirname) + string("\\") + string(ent->d_name);
            string dstFilename = string(ent->d_name);
            
            //check if it is a dir
            struct stat info;
            if(stat(srcFilename.c_str(), &info) != 0)
            {
                wrong_value = srcFilename;
                status = CANT_STAT_DIR;
                break;
            }
            if((info.st_mode & S_IFDIR) == 0)
            {
                //file: copy
                copyFile(srcFilename, dstFilename);
                if(status != NO_ERROR)
                    break;
            }
            else
            {
                //subdir: create folder, copy contents
                int32_t dirStatus = 0;
#if defined (_WIN32) || defined (__CYGWIN__)
                dirStatus = mkdir(ent->d_name);
#else
                dirStatus = mkdir(ent->d_name, 0777);
#endif
                if((dirStatus == -1) && (errno != EEXIST))
                {
                    wrong_value = string(ent->d_name);
                    status = CANT_CREATE_NEW_DIR;
                    break;
                }
                if(chdir(ent->d_name) != 0)
                {
                    wrong_value = string(ent->d_name);
                    status = CANT_CHANGE_DIR;
                    break;
                }
                copyFilesInDir(srcFilename.c_str());
                if(chdir("..") != 0)
                {
                    wrong_value = string(dirname);
                    status = CANT_CHANGE_DIR;
                    break;
                }
            }
        }
        closedir (dir);
    }
}

void mde_to_ooxml::copyFile(string srcFilename, string dstFilename)
{
    char buf[BUFFER_SIZE];
    size_t size;
    int src = -1, dst = -1;
    
    if((src = open(srcFilename.c_str(), O_RDONLY | O_BINARY, 0)) == -1)
    {
        status = CANT_OPEN_FILE_TO_COPY_FROM;
        wrong_value = srcFilename;
        return;
    }
    if((dst = open(dstFilename.c_str(), O_WRONLY | O_CREAT | O_TRUNC | O_BINARY, 0644)) == -1)
    {
        close(src);
        status = CANT_OPEN_FILE_TO_COPY_TO;
        wrong_value = dstFilename;
        return;
    }

    while((size = read(src, buf, BUFFER_SIZE)) > 0) {
        write(dst, buf, size);
    }
    close(src);
    close(dst);
    if(size < 0)
    {
        status = ERROR_WHILE_COPYING;
        wrong_value = srcFilename;
        return;
    }
}

string mde_to_ooxml::getAlignmentOOXML(const char * alignment_mde)
{
    string ret;
    if(!strncmp(alignment_mde, "justified", 9))
        ret = "\t\t<w:jc w:val=\"both\" />";
    else
        ret = "\t\t<w:jc w:val=\"" + string(alignment_mde) + "\"/>";
    return ret;
}

string mde_to_ooxml::getStyleOOXML(const char * style_mde, bool isTable)
{
    string ret;
    if(isTable)
        ret = "\t\t<w:tblStyle w:val=\"" + string(style_mde) + "\"/>";
    else
        ret = "\t\t<w:pStyle w:val=\"" + string(style_mde) + "\"/>";
    return ret;
}

string mde_to_ooxml::getIndentOOXML(const char * indent_mde, bool isTable)
{
    ostringstream ret_ss;
    if(isTable)
        ret_ss << "\t\t<w:tblInd w:w=\"" << ceil(stod(removeQuotes(indent_mde))*TWIPS_PER_CM) << "\" w:type=\"dxa\"/>";
    else
        ret_ss << "\t\t<w:ind w:left=\"" << ceil(stod(removeQuotes(indent_mde))*TWIPS_PER_CM) << "\"/>";
    return ret_ss.str();
}

string mde_to_ooxml::sanitize(string str)
{
    string ret(str);
    replaceString(ret, "Â°", "°");
    replaceString(ret, "â€“", "–");
#if 0
    replaceString(ret, "&", "&amp;");
    replaceString(ret, "<", "&lt;");
    replaceString(ret, ">", "&gt;");
    replaceString(ret, "\"", "&quot;");
#endif
    replaceString(ret, "\'", "&apos;");
    return ret;
}

string mde_to_ooxml::removeQuotes(const char * str)
{
    string ret(str);
    replaceString(ret, "\"", "");
    return ret;
}

void mde_to_ooxml::replaceString(string &str, const string& from,
        const string& to)
{
    size_t pos = 0;
    while((pos = str.find(from, pos)) != string::npos) {
        str.replace(pos, from.length(), to);
        pos += to.length(); // Handles case where 'to' is a substring of 'from'
    }
}

void mde_to_ooxml::zipFile(const char *filename)
{
    FILE * fp;

    fp = fopen(filename, "rb");
    if(fp == NULL)
    {
        wrong_value = string(filename);
        status = ZIPPING_ERROR;
        return;
    }
    
    unsigned char buf[BUFFER_SIZE];
    size_t filelen = 0, lenRead = 0;
    
    fseek(fp, 0, SEEK_END);
    filelen = ftell(fp);
    rewind(fp);
    
    if(zipOpenNewFileInZip64(zipfp, filename, NULL, NULL, 0, NULL, 0, NULL,
            Z_DEFLATED, Z_DEFAULT_COMPRESSION, (filelen > 0xffffffff)?1:0) != ZIP_OK)
    {
        fclose(fp);
        wrong_value = string(filename);
        status = ZIPPING_ERROR;
        return;
    }

    while((lenRead = fread(buf, sizeof(*buf), sizeof(buf), fp)) > 0)
    {
        if(zipWriteInFileInZip(zipfp, buf, lenRead) != ZIP_OK)
        {
            fclose(fp);
            zipCloseFileInZip(zipfp);
            wrong_value = string(filename);
            status = ZIPPING_ERROR;
            return;
        }
    }
    zipCloseFileInZip(zipfp);
    fclose(fp);
}

void mde_to_ooxml::zipFilesInDir(const char * dirname, bool addFullPath)
{
    DIR *dir;
    struct dirent *ent;
    if((dir = opendir(dirname)) == NULL)
    {
        wrong_value = string(dirname);
        status = ZIPPING_ERROR;
        return;
    }
    else
    {
        while((ent = readdir(dir)) != NULL)
        {
            if(((strlen(ent->d_name) == 1) && (strncmp(ent->d_name, ".", 1) == 0)) ||
                    ((strlen(ent->d_name) == 2) && (strncmp(ent->d_name, "..", 2) == 0)) ||
                    (strncmp(&ent->d_name[strlen(ent->d_name)-4], ".zip", 4) == 0))
            {
                continue;
            }
            string srcFilename;
            if(addFullPath)
                srcFilename = string(dirname) + string("\\") + string(ent->d_name);
            else
                srcFilename = string(ent->d_name);
            
            //check if it is a dir
            struct stat info;
            if(stat(srcFilename.c_str(), &info) != 0)
            {
                wrong_value = srcFilename;
                status = CANT_STAT_DIR;
                break;
            }
            if((info.st_mode & S_IFDIR) == 0)
            {
                //file: zip
                zipFile(srcFilename.c_str());
                if(status != NO_ERROR)
                    break;
            }
            else
            {
                //subdir: check for files
                zipFilesInDir(srcFilename.c_str(), true);
                if(status != NO_ERROR)
                    break;
            }
        }
        closedir (dir);
    }
}

void mde_to_ooxml::zipEmptyDir(const char *dirname)
{
    if((zipfp == NULL) || (dirname == NULL) || (*dirname == '\0'))
    {
        wrong_value = string(dirname);
        status = ZIPPING_ERROR;
        return;
    }
    size_t dirnamelen = strlen(dirname);
    char * temp = (char*)calloc(1, dirnamelen+2);
    
    memcpy(temp, dirname, dirnamelen+2);
    if(temp[dirnamelen-1] != '/')
    {
        temp[dirnamelen] = '/';
        temp[dirnamelen+1] = '\0';
    }
    else
    {
        temp[dirnamelen] = '\0';
    }

    if(zipOpenNewFileInZip64(zipfp, temp, NULL, NULL, 0, NULL, 0, NULL, 0, 0, 0)
            !=  ZIP_OK)
    {
        wrong_value = string(dirname);
        status = ZIPPING_ERROR;
        return;
    }
    free(temp);
    zipCloseFileInZip(zipfp);
}

void mde_to_ooxml::removeDir(const char *dirname)
{
    DIR *dir;
    struct dirent *ent;
    if((dir = opendir(dirname)) == NULL)
    {
        wrong_value = string(dirname);
        status = REMOVE_ERROR;
        return;
    }
    else
    {
        while((ent = readdir(dir)) != NULL)
        {
            if(((strlen(ent->d_name) == 1) && (strncmp(ent->d_name, ".", 1) == 0)) ||
                    ((strlen(ent->d_name) == 2) && (strncmp(ent->d_name, "..", 2) == 0)))
            {
                continue;
            }
            string srcFilename = string(dirname) + string("\\") + string(ent->d_name);
            
            //check if it is a dir
            struct stat info;
            if(stat(srcFilename.c_str(), &info) != 0)
            {
                wrong_value = srcFilename;
                status = CANT_STAT_DIR;
                break;
            }
            if((info.st_mode & S_IFDIR) == 0)
            {
                //file: remove
                remove(srcFilename.c_str());
            }
            else
            {
                //subdir: check for files
                removeDir(srcFilename.c_str());
                rmdir(srcFilename.c_str());
            }
        }
        closedir (dir);
        rmdir(dirname);
    }
}