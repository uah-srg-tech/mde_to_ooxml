/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/* 
 * File:   mde_to_ooxml.h
 * Author: user
 *
 * Created on 16 de septiembre de 2019, 13:57
 */

#include <cstdint>
#include <string>
#include <fstream>
#include <vector>
#include "../gss_qt/XMLParsingTools/XMLTools.h"

#ifndef MDE_TO_OOXML_H
#define MDE_TO_OOXML_H

#define ZIP

class mde_to_ooxml {
public:
    mde_to_ooxml(const char * config_filename_c_str);
    uint32_t createDOCXfromMDE(std::string& ddoc_filename);
    std::string displayErrorCreateDOCXfromMDE();
    
private:
    enum Status : uint32_t {
        NO_ERROR,
        CANT_OPEN_CONFIG_FILE,
        CONFIG_FILE_NOT_VALID,
        CANT_CREATE_OOXML_FILE,
        CANT_OPEN_OOXML_FILE,
        CANT_STAT_DIR,
        FOLDER_IS_NOT_A_DIR,
        CANT_CREATE_NEW_DIR,
        CANT_CHANGE_DIR,
        CANT_OPENDIR_DIR,
        FILE_NOT_FOUND,
        FILE_PARSING_ERROR,
        WRONG_TAG,
        XML_PARSING_ERROR,
        ZIPPING_ERROR,
        RENAME_ERROR,
        REMOVE_ERROR,
        CANT_OPEN_FILE_TO_COPY_FROM,
        CANT_OPEN_FILE_TO_COPY_TO,
        ERROR_WHILE_COPYING
    };
    
    enum FromFileType : uint32_t {
        FIGURE,
        TABLE
    };
    
    enum ListContent : uint32_t {
        ENUMERATE,
        ITEMIZE
    };
    
    void parseMDE(std::string ddoc_filename);
    void parseSection(xmlNodePtr sectionHandle, uint32_t currentLevel,
            uint32_t parentCurrentSection, uint32_t parentSections);
    void parseBody(xmlNodePtr bodyHandle, std::string tab);
    void parseParagraph(xmlNodePtr paragraphHandle, std::string tab,
             const std::string &pPr = std::string());
    void parseRun(xmlNodePtr runHandle, std::string tab);
    void parseFromFile(xmlNodePtr fromFileHandle, std::string tab,
            FromFileType type);
    void parseTable(xmlNodePtr tableHandle, std::string tab);
    void parseListContent(xmlNodePtr listHandle, std::string tab,
            uint32_t currentListLevel, ListContent listContent);
    void parseCaption(const char * captionText, FromFileType type,
            std::string tab, std::string alignment);
    
    void removeLastLine(std::string filename);
    void addLastLine(std::string filename, std::string lineToAdd);
    void copyFilesInDir(const char * dirname);
    void copyFile(std::string srcFilename, std::string dstFilename);
    
    std::string getAlignmentOOXML(const char * alignment_mde);
    std::string getStyleOOXML(const char * style_mde, bool isTable);
    std::string getIndentOOXML(const char * indent_mde, bool isTable);
    
    std::string sanitize(std::string str);
    std::string removeQuotes(const char * str);
    void replaceString(std::string &str, const std::string& from, const std::string& to);
    
    void zipFile(const char *filename);
    void zipFilesInDir(const char *filename, bool addFullPath);
    void zipEmptyDir(const char *dirname);
    
    void removeDir(const char *dirname);
    
    void* zipfp;//zipFile zipfp;

    Status status;
    int32_t xmlParsingStatus;
    
    std::string config_filename;
    
    std::string ddoc_file_location;
    std::string ddoc_file_name;
    std::string ooxml_location;
    std::ofstream ooxml_file;
    std::ofstream ooxml_aux_file;
    
    std::string errorSection;
    std::string wrong_value;
    std::string wrong_value_expected;
    
    uint32_t bookmarks;
    uint32_t relationships;
    uint32_t numIds;
    uint32_t figures;
    uint32_t tables;
    
    xmlDocPtr xmlDoc;
    
    static const uint32_t MAX_AUX_STRING = 3072;
    char auxString[MAX_AUX_STRING];
    
    static const size_t BUFFER_SIZE = 4096;
    
    static const uint32_t TWIPS_PER_CM = 567;
    static const uint32_t EMU_PER_INCH = 914400;
    static const uint32_t PIXEL_PER_INCH = 150;
    static const uint32_t EMU_PER_PIXEL = EMU_PER_INCH/PIXEL_PER_INCH;
    static const uint32_t MAX_EMU_WIDTH = 5400000;
    static const uint32_t MAX_EMU_HEIGTH = 8820000;
};
#endif /* MDE_TO_OOXML_H */

