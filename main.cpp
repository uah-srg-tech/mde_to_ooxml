/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/* 
 * File:   main.cpp
 * Author: user
 *
 * Created on 30 de julio de 2019, 18:04
 */

#include <iostream>
#include "mde_to_ooxml.h"

using namespace std;

/*
 * 
 */
int main(int argc, char** argv)
{
    int32_t status = 0;
    string ddocFile;
    
    mde_to_ooxml parser(argv[1]);
    
    status = parser.createDOCXfromMDE(ddocFile);
    if(!ddocFile.empty())
    {
        cout << endl << "Using \"" << ddocFile << "\"" << endl;
    }
    if(status != 0)
    {
        cout << parser.displayErrorCreateDOCXfromMDE() << endl;
        status = -1;
    }
    else
    {
        cout << "MDE to OOXML OK" << endl;
    }
    getchar();
    return status;
}