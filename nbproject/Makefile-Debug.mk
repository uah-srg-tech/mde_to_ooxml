#
# Generated Makefile - do not edit!
#
# Edit the Makefile in the project folder instead (../Makefile). Each target
# has a -pre and a -post target defined where you can add customized code.
#
# This makefile implements configuration specific macros and targets.


# Environment
MKDIR=mkdir
CP=cp
GREP=grep
NM=nm
CCADMIN=CCadmin
RANLIB=ranlib
CC=gcc
CCC=g++
CXX=g++
FC=gfortran
AS=as

# Macros
CND_PLATFORM=MinGW-Windows
CND_DLIB_EXT=dll
CND_CONF=Debug
CND_DISTDIR=dist
CND_BUILDDIR=build

# Include project Makefile
include Makefile

# Object Directory
OBJECTDIR=${CND_BUILDDIR}/${CND_CONF}/${CND_PLATFORM}

# Object Files
OBJECTFILES= \
	${OBJECTDIR}/_ext/6e2d5a1d/XMLTools.o \
	${OBJECTDIR}/main.o \
	${OBJECTDIR}/mde_to_ooxml.o


# C Compiler Flags
CFLAGS=

# CC Compiler Flags
CCFLAGS=
CXXFLAGS=

# Fortran Compiler Flags
FFLAGS=

# Assembler Flags
ASFLAGS=

# Link Libraries and Options
LDLIBSOPTIONS=-lxml2 -liconv -lminizip -lz

# Build Targets
.build-conf: ${BUILD_SUBPROJECTS}
	"${MAKE}"  -f nbproject/Makefile-${CND_CONF}.mk ${CND_CONF}/mde_to_ooxml.exe

${CND_CONF}/mde_to_ooxml.exe: ${OBJECTFILES}
	${MKDIR} -p ${CND_CONF}
	${LINK.cc} -o ${CND_CONF}/mde_to_ooxml ${OBJECTFILES} ${LDLIBSOPTIONS}

${OBJECTDIR}/_ext/6e2d5a1d/XMLTools.o: ../massiva/XMLParsingTools/XMLTools.c
	${MKDIR} -p ${OBJECTDIR}/_ext/6e2d5a1d
	${RM} "$@.d"
	$(COMPILE.c) -g -Wall -DLIBXML_STATIC -I/C/msys/1.0/mingw32/include/libxml2 -I/C/msys/1.0/mingw32/include/minizip -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/_ext/6e2d5a1d/XMLTools.o ../massiva/XMLParsingTools/XMLTools.c

${OBJECTDIR}/main.o: main.cpp
	${MKDIR} -p ${OBJECTDIR}
	${RM} "$@.d"
	$(COMPILE.cc) -g -Wall -DLIBXML_STATIC -I/C/msys/1.0/mingw32/include/libxml2 -I/C/msys/1.0/mingw32/include/minizip -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/main.o main.cpp

${OBJECTDIR}/mde_to_ooxml.o: mde_to_ooxml.cpp
	${MKDIR} -p ${OBJECTDIR}
	${RM} "$@.d"
	$(COMPILE.cc) -g -Wall -DLIBXML_STATIC -I/C/msys/1.0/mingw32/include/libxml2 -I/C/msys/1.0/mingw32/include/minizip -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/mde_to_ooxml.o mde_to_ooxml.cpp

# Subprojects
.build-subprojects:

# Clean Targets
.clean-conf: ${CLEAN_SUBPROJECTS}
	${RM} -r ${CND_BUILDDIR}/${CND_CONF}

# Subprojects
.clean-subprojects:

# Enable dependency checking
.dep.inc: .depcheck-impl

include .dep.inc
