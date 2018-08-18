#!/usr/bin/python
# -*-coding:utf-8-*-

import sys

from pdfminer.layout import *
from pdfminer.pdfparser import *
from pdfminer.pdfinterp import *
from pdfminer.converter import *


class PDFToWord:
    tmp = 'Not complete'


# TableConstructInfo
#
class TableConstructInfo:
    """
    store the info of table constructed
    """

    def __init__(self):
        self.colNum = 0  # col num
        self.rowNum = 0  # row num
        self.lineWidth = sys.float_info.max  #
        self.textBoxInTable = []  # text and their coorparation
        self.rectInTable = []  # rect extracted from table, type:LTRect
        self.text_info = []  # record info in each table cell
        self.allAbscissa = []  # all abscissa of rect in rectInTable
        self.allOrdinate = []  # all ordinate of rect in rectInTable
        #  locationOfAllRowRect : type: list(dict)
        #  list(dict) : locationOfEachRowRect index of list represent rowIndex of table,
        # dict['row'] : store whether the index of yo of horizontal line in allOrdinate exist
        # dict['col'] : store whether the index of x0 of vertical line in allAbscissa exist
        self.locationsOfAllLine = []
        # tableSope : left_bottom coordinate and right_top coordinate
        self.tableScope = {'xmin': 0, 'xmax': 0, 'ymin': 0, 'ymax': 0}
        # mergeInfo :type list(list), record the info of mergedUnit
        # second list represent a merged cell contains primary cell
        # cell represent with a int value : rowindex * self.colNum + colIndex
        self.mergedInfo = []

    def is_value_in_scope(self, val, scope):
        """
        is_value_in_scope
        :param val:
        :param scope: type : list , list[0] is mininum , list[1] is max
        :return:
        """
        allowError = self.lineWidth * 2
        if val > scope[0] - allowError and val < scope[1] + allowError:
            return True
        else:
            return False

    def calculateApproximateScale(self, oneNum, anotherNum):
        """
        the line got from pdf is not monotonous regulary , so we can only calculate approximate int value
        now temporary allow 1% error
        :param oneNum:
        :param anotherNum:
        :return: aprroximate scale : int
        """
        return int(oneNum / anotherNum + 0.01)

    def isApproximateEquals(self, oneNum, anotherNum, allowError):
        """
        judge is approximate equal the two number
        :param oneNum:
        :param anotherNum:
        :param allowError:
        :return:
        """
        return True if (
                oneNum - anotherNum > - allowError and oneNum - anotherNum < allowError) else False

    def getAllRowColCoordinate(self):
        """
        get a rowList store all row coordinate
        get a colList store all col coordinate
        use rowList and colList to get rowNum and colNum of table ,
        and each minimal cell's (assume that there's no merged cell)
        height and width
        """
        rowSet = ApproximateSet(self.lineWidth * 2)
        colSet = ApproximateSet(self.lineWidth * 2)
        for rect in self.rectInTable:
            rowSet.add(rect.x0)
            rowSet.add(rect.x1)
            colSet.add(rect.y0)
            colSet.add(rect.y1)

        self.allAbscissa = rowSet.getValue()
        self.allOrdinate = colSet.getValue()

    def getRowColNum(self):
        self.rowNum = len(self.allOrdinate) - 1
        self.colNum = len(self.allAbscissa) - 1

    def getIndexOfValInList(self, val, givenList):
        beginIndex = 0
        endindex = len(givenList) - 1
        allowError = self.lineWidth * 2
        while (beginIndex < endindex):
            mid = (int)((beginIndex + endindex) / 2)
            if (val >= givenList[mid] - allowError
                    and val <= givenList[mid] + allowError):
                return mid
            if (val > givenList[mid] + allowError):
                beginIndex = mid + 1
            if (val < givenList[mid] - allowError):
                endindex = mid - 1

        return beginIndex if self.isApproximateEquals(val,
                                                      givenList[beginIndex],
                                                      allowError) else -1

    def initLocationsOfAllLine(self):
        self.locationsOfAllLine = [] # range(self.rowNum + 1)
        for i in range(self.rowNum + 1):
            self.locationsOfAllLine.append({})
        for rowIndex in range(len(self.locationsOfAllLine)):
            isExistRowLine = []
            isExistColLine = []
            for i in range(self.colNum):
                isExistColLine.append(False)
                isExistRowLine.append(False)
            isExistColLine.append(False)
            # isExistRowLine = range(self.colNum)
            # isExistColLine = range(self.colNum + 1)
            # for index in range(len(isExistRowLine)):
            #     isExistRowLine[index] = False
            # for index in range(len(isExistColLine)):
            #     isExistColLine[index] = False

            self.locationsOfAllLine[rowIndex] = {'row': isExistRowLine,
                                                 'col': isExistColLine}

    def getAllLocationOfAllLine(self):
        self.initLocationsOfAllLine()
        for rect in self.rectInTable:
            rowIndex = self.getIndexOfValInList(rect.y1, self.allOrdinate) - 1
            if rowIndex > -1:
                if self.calculateApproximateScale(rect.height, rect.width) > 1:
                    rowIndex = self.rowNum - 1 - rowIndex
                    # rowIndex = self.rowNum - 1 - self.getIndexOfValInList(rect.y1, self.allOrdinate)
                    self.locationsOfAllLine[rowIndex]['col'][
                        self.getIndexOfValInList(rect.x1,
                                                 self.allAbscissa)] = True
                if self.calculateApproximateScale(rect.height, rect.width) < 1:
                    rowIndex = self.rowNum - 1 - rowIndex
                    # rowIndex = self.rowNum - 1 - self.getIndexOfValInList(rect.y1, self.allOrdinate)
                    self.locationsOfAllLine[rowIndex]['row'][
                        self.getIndexOfValInList(rect.x0,
                                                 self.allAbscissa)] = True

    def getMergedCellContainCurCell(self, curCell):
        for cells in self.mergedInfo:
            if (self.getIndexOfValInList(curCell, cells)) > 0:
                return cells

    def addMergeInfo(self, oldCell, newCell):
        if (oldCell < 0):
            return
        mergedCell = self.getMergedCellContainCurCell(oldCell)
        if (mergedCell == None):
            self.mergedInfo.append([oldCell, newCell])
        else:
            if (newCell - oldCell == 1):
                mergedCell.append(newCell)
            else:
                lenOfMerged = len(mergedCell)
                for i in range(lenOfMerged - 1):
                    mergedCell.append(newCell + i)

    def getMergedCellInfo(self):
        """
        getMergedInfo: mergedCell is a list(int),
        a int value in list represent a cell :
        val = rowIndex * self.colNum + colIndex
        """
        traverseIndexInRectInTable = 0
        for rowIndex in range(self.rowNum):
            rowsExistInfo = self.locationsOfAllLine[rowIndex]['row']
            colsExistInfo = self.locationsOfAllLine[rowIndex]['col']
            for colIndex in range(self.colNum):
                if (rowsExistInfo[colIndex] == False):
                    self.addMergeInfo((rowIndex - 1) * self.colNum + colIndex,
                                      rowIndex * self.colNum + colIndex)

            for colIndex in range(self.colNum):
                if (colsExistInfo[colIndex]) == False:
                    self.addMergeInfo(rowIndex * self.colNum + colIndex - 1,
                                      rowIndex * self.colNum + colIndex)

    def __init__text_info(self):
        self.text_info = []
        for index in range(self.colNum * self.rowNum):
            self.text_info.append("")

    def interprete_tab_info(self):
        """
        interprete Tab Info
        :return:
        """
        self.getAllRowColCoordinate()
        self.getRowColNum()
        self.getAllLocationOfAllLine()
        self.getMergedCellInfo()
        self.__init__text_info()

    def get_val_in_which_segment_in_list(self, val, list):
        """
        get_val_in_which_segment_in_list :
        if  list[index]<=val<=list[index+1] , we return index,
        :return:
        """
        headIndex = 0
        tailIndex = len(list) - 1
        allowError = self.lineWidth * 2
        while (headIndex < tailIndex - 1):
            mid = int((headIndex + tailIndex) / 2)
            if self.isApproximateEquals(val, list[mid], self.lineWidth * 2):
                return mid
            if val > list[mid] + allowError:
                headIndex = mid
            if val < list[mid] - allowError:
                tailIndex = mid

        return headIndex

    def cell_filter_with_merged_info(self, cell_list):
        """
        filter cell belong to one merged cell to leave one
        :return: cells belong to different cell(merged cell as a cell)
        """
        result = []
        tmp_result = []
        if len(self.mergedInfo) == 0:
            return cell_list
        for merged_cell in self.mergedInfo:
            has_cell_in_merged_cell = False
            index_merged_cell = 0
            index_cell_list = 0
            while (index_cell_list < len(cell_list) and index_merged_cell < len(
                    merged_cell)):
                while (index_cell_list < len(cell_list)):
                    if merged_cell[index_merged_cell] == cell_list[
                        index_cell_list]:
                        if has_cell_in_merged_cell == False:
                            result.append(cell_list[index_cell_list])
                            has_cell_in_merged_cell = True

                        index_merged_cell += 1
                        index_cell_list += 1
                        break
                    if merged_cell[index_merged_cell] > cell_list[
                        index_cell_list]:
                        tmp_result.append(cell_list[index_cell_list])
                        index_merged_cell += 1
                        index_cell_list += 1
                        break
                    if merged_cell[index_merged_cell] < cell_list[
                        index_cell_list]:
                        index_merged_cell += 1
                        break

            cell_list = tmp_result
            tmp_result = []

        for str in tmp_result:
            result.append(str)

        return result

    def get_cell_overlap_with_text_box(self, text_box):
        x0_index = self.get_val_in_which_segment_in_list(text_box.x0,
                                                         self.allAbscissa)
        x1_index = self.get_val_in_which_segment_in_list(text_box.x1,
                                                         self.allAbscissa)
        y0_index = self.rowNum - 1 - self.get_val_in_which_segment_in_list(
            text_box.y0,
            self.allOrdinate)
        y1_index = self.rowNum - 1 - self.get_val_in_which_segment_in_list(
            text_box.y1,
            self.allOrdinate)

        cell_list = []

        for row_index in range(y1_index, y0_index + 1):
            for col_index in range(x0_index, x1_index + 1):
                cell_list.append(row_index * self.colNum + col_index)

        return cell_list

    def filter_empty_str(self, str_list):
        """
        filter_empty_str
        :param str_list:
        :return: filtered_list
        """
        result = []
        for str in str_list:
            if (str != ""):
                result.append(str)

        return result

    def add_text_box(self, text_box):
        """
        add_text_box
        :param text_box: type : LTTextBox
        :return:
        """
        cells_overlap = self.get_cell_overlap_with_text_box(text_box)
        if len(cells_overlap) == 1:
            self.text_info[cells_overlap[0]] += text_box.get_text()
            return
        else:
            cells_overlap = self.cell_filter_with_merged_info(cells_overlap)

            if len(cells_overlap) == 1:
                self.text_info[cells_overlap[0]] += text_box.get_text()
                return

            text_list = self.filter_empty_str(text_box.get_text().split(" "))
            if (len(text_list) <= len(cells_overlap)):
                for index in range(len(text_list)):
                    self.text_info[cells_overlap[index]] += text_list[index]
                return
            # the number cells_overlap less than the number of text_list,
            # Extra text is put last temporarily
            if (len(text_list) > len(cells_overlap)):
                for index in range(len(cells_overlap)):
                    if (index == len(cells_overlap) - 1):
                        str = ""
                        for text in text_list[index:]:
                            str += text
                        self.text_info[cells_overlap[index]] += str
                        return

                    self.text_info[cells_overlap[index]] += text_list[index]

            return


# ApproximateSet
#
class ApproximateSet:
    """
    implements a Set dealing with approximate value
    add funtion : return sorted result
    """

    def __init__(self, allowRrror=0):
        self.valList = []
        self.allowError = allowRrror

    def calculateApproximateScale(self, oneNum, anotherNum):
        """
        the line got from pdf is not monotonous regulary , so we can only calculate approximate int value
        now temporary allow 1% error
        :param oneNum:
        :param anotherNum:
        :return: aprroximate scale : int
        """
        return int(oneNum / anotherNum + 0.01)

    def isApproximateEquals(self, oneNum, anotherNum, allowError):
        """
        judge is approximate equal the two number
        :param oneNum:
        :param anotherNum:
        :param allowError:
        :return:
        """
        return True if (
                oneNum - anotherNum > - allowError and oneNum - anotherNum < allowError) else False

    def add(self, valAdded):
        for index in range(len(self.valList)):
            if self.isApproximateEquals(valAdded, self.valList[index],
                                        self.allowError):
                return
            if valAdded < self.valList[index] - self.allowError:
                tmp = self.valList[index]
                self.valList[index] = valAdded
                valAdded = tmp

        self.valList.append(valAdded)

    def getValue(self):
        return self.valList


# PageInfo
#
class PageInfo:
    """store info of a page"""

    def __init__(self):
        # extractInfoDic = dic:{'text': list(LTTextBox), 'rect': list(LTRect)}
        self.extractInfoDic = {}
        # tableList list(TableConstructInfo)
        self.tableList = []
        # store dataInfo : dataInfo contain a entire table and a whole paragraph
        # and keep object order
        self.dataObjects = []

    def isCanMergeLTRect(self, oneRect, anotherRect):
        """"
        :type oneRect : LTRect
        :type anotherRect : LTRect
        :rtype boolean
        """
        # case one : oneRect left side and anotherRect right side overlap each other
        # case two : oneRect right side and anotherRect left side overlap each other
        # summary : left or right side overlap each other,their y0==y0 and y1==y1
        if (oneRect.y0 == anotherRect.y0 and oneRect.y1 == anotherRect.y1 and (
                oneRect.x0 == anotherRect.x1 or oneRect.x1 == anotherRect.x0)):
            return True
        # case three : oneRect top and anotherRect bottom overlap each other
        # case four : oneRect bottom and anotherRect top overlap each other
        # summary : top or bottom overlap each other ,their x0==x0 and x1==x1
        if (oneRect.x0 == anotherRect.x0 and oneRect.x1 == anotherRect.x1 and (
                oneRect.y0 == anotherRect.y1 or oneRect.y1 == anotherRect.y0)):
            return True
        return False

    def mergeRect(self, oneRect, anotherRect):
        """
        :type oneRect : LTRect
        :type anotherRect : LTRect
        :rtype LTRect
        """
        if self.isCanMergeLTRect(oneRect, anotherRect):
            # case one : oneRect is in the top left
            if (oneRect.x1 == anotherRect.x0 or oneRect.y1 == anotherRect.y0):
                return LTRect(0, (
                oneRect.x0, oneRect.y0, anotherRect.x1, anotherRect.y1))
            # case two : oneRect is in the right bottom
            else:
                return LTRect(0, (
                anotherRect.x0, anotherRect.y0, oneRect.x1, oneRect.y1))
        else:
            raise Exception("Can not merge!")

    def calculateApproximateScale(self, oneNum, anotherNum):
        """
        the line got from pdf is not monotonous regulary , so we can only calculate approximate int value
        now temporary allow 1% error
        :param oneNum:
        :param anotherNum:
        :return: aprroximate scale
        """
        return int(oneNum / anotherNum + 0.01)

    def isApproximateEquals(self, oneNum, anotherNum, allowError):
        """
        judge is approximate equal the two number
        :param oneNum:
        :param anotherNum:
        :param allowError:
        :return:
        """
        return True if (
                oneNum - anotherNum > - allowError and oneNum - anotherNum < allowError) else False

    def filtrateLTRect(self):
        """
        filtrate unnecessary line by merging
        :return:
        """
        filtratedRect = []
        needMergedRect = []
        for rect in self.extractInfoDic['rect']:
            if (self.calculateApproximateScale(rect.height, rect.width) == 1):
                if (len(filtratedRect) > 0):
                    curLastRect = filtratedRect.pop()
                    if self.isCanMergeLTRect(curLastRect, rect):
                        filtratedRect.append(self.mergeRect(curLastRect, rect))
                    else:
                        filtratedRect.append(curLastRect)
                        needMergedRect.append(rect)
                else:
                    needMergedRect.append(rect)
            else:
                if (len(needMergedRect) == 0):
                    filtratedRect.append(rect)
                else:
                    curLastRect = needMergedRect.pop();
                    if self.isCanMergeLTRect(curLastRect, rect):
                        filtratedRect.append(self.mergeRect(rect, curLastRect))
                    else:
                        filtratedRect.append(rect)
                        needMergedRect.append(curLastRect)

        self.extractInfoDic['rect'] = filtratedRect

    def divideRectIntoDeffrentTable(self):
        """
        divideRectIntoDeffrentTable
        :return:
        """
        for rect in self.extractInfoDic['rect']:
            if len(self.tableList) == 0:
                tab = TableConstructInfo()
                tab.rectInTable.append(rect)
                tab.tableScope['xmin'] = rect.x0
                tab.tableScope['xmax'] = rect.x1
                tab.tableScope['ymin'] = rect.y0
                tab.tableScope['ymax'] = rect.y1
                tab.lineWidth = min(tab.lineWidth,
                                    rect.y1 - rect.y0,
                                    rect.x1 - rect.x0)
                self.tableList.append(tab)
            else:
                tab = self.getTabContainsThisRect(rect)
                if tab == None:
                    tab = TableConstructInfo()
                    tab.rectInTable.append(rect)
                    tab.tableScope['xmin'] = rect.x0
                    tab.tableScope['xmax'] = rect.x1
                    tab.tableScope['ymin'] = rect.y0
                    tab.tableScope['ymax'] = rect.y1
                    tab.lineWidth = min(tab.lineWidth,
                                        rect.y1 - rect.y0,
                                        rect.x1 - rect.x0)
                    self.tableList.append(tab)
                else:
                    tab.rectInTable.append(rect)
                    tab.tableScope['xmax'] = max(rect.x1,
                                                 tab.tableScope['xmax'])
                    tab.tableScope['ymin'] = min(rect.y0,
                                                 tab.tableScope['ymin'])
                    tab.lineWidth = min(tab.lineWidth,
                                        rect.y1 - rect.y0,
                                        rect.x1 - rect.x0)

    def getTabContainsThisRect(self, rect):
        """
        getTabContainsThisRect
        :type rect LTRect
        :return: TableConstructInfo
        """
        for tab in self.tableList:
            allowRrror = 2 * tab.lineWidth
            # case one : horizontal line
            if self.calculateApproximateScale(rect.height, rect.width) == 0:
                if rect.x0 > tab.tableScope['xmin'] - allowRrror \
                        and rect.x0 < tab.tableScope['xmax'] + allowRrror \
                        and rect.y1 > tab.tableScope['ymin'] - allowRrror \
                        and rect.y1 < tab.tableScope['ymax'] + allowRrror:
                    return tab
            # case two : vertical line
            else:
                if rect.x1 > tab.tableScope['xmin'] - allowRrror \
                        and rect.x1 < tab.tableScope['xmax'] + allowRrror \
                        and rect.y1 > tab.tableScope['ymin'] - allowRrror \
                        and rect.y1 < tab.tableScope['ymax'] + allowRrror:
                    return tab

    def dealTabInfo(self):
        for tab in self.tableList:
            tab.interprete_tab_info()

    def getTabContainTextBox(self, textBox):
        """
        getTabContainTextBox
        :param textBox: type: LTTextBox
        :return : tabInfo
        """
        for tab in self.tableList:
            if (tab.is_value_in_scope(textBox.x0, [tab.tableScope['xmin'],
                                                   tab.tableScope['xmax']])) \
                    and (tab.is_value_in_scope(textBox.x1,
                                               [tab.tableScope['xmin'],
                                                tab.tableScope['xmax']])) \
                    and (tab.is_value_in_scope(textBox.y0,
                                               [tab.tableScope['ymin'],
                                                tab.tableScope['ymax']])) \
                    and (tab.is_value_in_scope(textBox.y1,
                                               [tab.tableScope['ymin'],
                                                tab.tableScope['ymax']])):
                return tab

    def isExistTheTabInDataObject(self, tab):
        """
        isExistTheTabInDataObject
        :param tab: type : TableConstructInfo
        :return:
        """
        for object in self.dataObjects:
            if object == tab:
                return True
        return False

    def getDataObject(self):
        for textBox in self.extractInfoDic['text']:
            tab = self.getTabContainTextBox(textBox)
            if (tab == None):
                self.dataObjects.append(textBox.get_text())
            else:
                tab.add_text_box(textBox)
                if (self.isExistTheTabInDataObject(tab) == False):
                    self.dataObjects.append(tab)

    def interpretePageInfo(self):
        self.filtrateLTRect()
        self.divideRectIntoDeffrentTable()
        self.dealTabInfo()
        self.getDataObject()


# extract
#
class PDFExtract:
    # extractInfoDic = dic:{'text': list(LTTextBox), 'rect': list(LTRect)}
    # extractInfoDic = {}

    def __init__(self):
        # the page object in a pdf file
        self.pages = []

    def extractLTTextBoxAndLTRect(self, filePath):
        """
        extract the LTTextBox 、LTRect info from pdf
        :type filePath: str
        """
        op = open(filePath, 'rb')
        parser = PDFParser(op)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize()

        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in doc.get_pages():
            textBoxList = []
            rectList = []
            interpreter.process_page(page)
            layout = device.get_result()
            pageInfo = PageInfo()
            for x in layout:
                if isinstance(x, LTTextBox):
                    textBoxList.append(x)
                if isinstance(x, LTRect):
                    rectList.append(x)
            pageInfo.extractInfoDic = {'text': textBoxList, 'rect': rectList}
            pageInfo.interpretePageInfo()
            self.pages.append(pageInfo)

    # def getAllRowColCoordinate(self, rectList):
    #     """
    #     get a rowList store all row coordinate
    #     get a colList store all col coordinate
    #     use rowList and colList to get rowNum and colNum of table ,
    #     and each minimal cell's (assume that there's no merged cell)
    #     height and width
    #     :type rectList : list(LTRect) , rectList is all line of a table
    #     :rtype : rowList,colList : list(float)
    #     """
    #     rowSet = ApproximateSet()
    #     colSet = ApproximateSet()
    #     for rect in rectList:
    #         rowSet.add(rect.x0)
    #         rowSet.add(rect.x1)
    #         colSet.add(rect.y0)
    #         colSet.add(rect.y1)
    #
    #     rowList = rowSet.getValue()
    #     colList = colSet.getValue()
    #
    #     return rowList, colList
    #
    # def getRowColNum(self, table):
    #     return 1
    #
    # def getMergedCellInfo(self):
    #     """
    #
    #     :return:
    #     """
    #
    #     return 1

    def parse(self, path):
        return False

