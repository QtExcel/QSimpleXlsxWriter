#
# QSimpleXlsxWriter.pro
#

TARGET = Qlibxlsxwriter
TEMPLATE = lib
CONFIG += staticlib

DEFINES += QT_DEPRECATED_WARNINGS
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

# SIMPLE_XLSX_WRITER_PARENTPATH = ../simplexlsx-code/
include(./QSimpleXlsxWriter.pri)
