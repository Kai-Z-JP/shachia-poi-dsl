package jp.kaiz.shachia.dianet.dsl.poi

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.*

fun workbook(block: XSSFWorkbookBuilder.() -> Unit): XSSFWorkbook {
    val builder = XSSFWorkbookBuilder()
    builder.block()

    return builder.build()
}

class XSSFWorkbookBuilder {
    private val workbook = XSSFWorkbook()
    var defaultFontHeightInPooints: Short
        get() = workbook.getFontAt(0).fontHeightInPoints
        set(value) {
            workbook.getFontAt(0).fontHeightInPoints = value
        }
    var defaultFontName: String
        get() = workbook.getFontAt(0).fontName
        set(value) {
            workbook.getFontAt(0).fontName = value
        }

    fun sheet(name: String, block: XSSFSheetBuilder.() -> Unit) {
        val sheet = workbook.createSheet(name)
        val builder = XSSFSheetBuilder(sheet)
        builder.block()
        builder.build()
    }

    fun build(): XSSFWorkbook {
        return workbook
    }

    fun style(
        baseStyle: XSSFCellStyle? = null,
        vararg blocks: XSSFCellStyleBuilder.() -> Unit,
        block: XSSFCellStyleBuilder.() -> Unit = {}
    ): XSSFCellStyle {
        val builder = XSSFCellStyleBuilder()
        blocks.forEach { builder.it() }
        builder.block()
        return builder.build(workbook, baseStyle)
    }

    fun font(block: XSSFFontBuilder.() -> Unit): XSSFFont {
        val builder = XSSFFontBuilder()
        builder.block()
        return builder.build(workbook)
    }
}

class XSSFSheetBuilder(private val sheet: XSSFSheet) {
    var offsetRow = 0
    var offsetColumn = 0
    var zoom = 100
    var defaultRowHeightInPoints: Float
        get() = sheet.defaultRowHeightInPoints
        set(value) {
            sheet.defaultRowHeightInPoints = value
        }
    var defaultColumnWidth: Int
        get() = sheet.defaultColumnWidth
        set(value) {
            sheet.defaultColumnWidth = value
        }

    var repeatingColumns: CellRangeAddress?
        get() = sheet.repeatingColumns
        set(value) {
            sheet.repeatingColumns = value
        }

    private val applyLate = mutableListOf<() -> Unit>()

    fun row(block: XSSFRowBuilder.() -> Unit) {
        if (sheet.physicalNumberOfRows < offsetRow) {
            repeat(offsetRow - sheet.physicalNumberOfRows) {
                val row = sheet.createRow(sheet.physicalNumberOfRows)
                if (defaultRowHeightInPoints != 0f) {
                    row.heightInPoints = defaultRowHeightInPoints
                }
            }
        }

        val row = sheet.createRow(sheet.physicalNumberOfRows)

        if (defaultRowHeightInPoints != 0f) {
            row.heightInPoints = defaultRowHeightInPoints
        }

        val builder = XSSFRowBuilder(sheet, row)
        repeat(offsetColumn) { builder.cell {} }
        builder.block()
        builder.build()
    }

    fun createFreezePane(colSplit: Int, rowSplit: Int) {
        sheet.createFreezePane(colSplit, rowSplit)
    }

    fun createFreezePane(colSplit: Int, rowSplit: Int, leftmostColumn: Int, topRow: Int) {
        sheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow)
    }

    fun build() {
        sheet.setZoom(zoom)
        applyLate.forEach { it() }
    }

    fun setColumnWidth(columnIndex: Int, width: Int) {
        sheet.setColumnWidth(columnIndex, width)
    }

    fun setBorderTop(borderStyle: BorderStyle, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderTop(
                borderStyle,
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                sheet
            )
        }
    }

    fun setBorderBottom(borderStyle: BorderStyle, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderBottom(
                borderStyle,
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                sheet
            )
        }
    }

    fun setBorderLeft(borderStyle: BorderStyle, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderLeft(
                borderStyle,
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                sheet
            )
        }
    }

    fun setBorderRight(borderStyle: BorderStyle, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderRight(
                borderStyle,
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                sheet
            )
        }
    }
}

class XSSFRowBuilder(private val sheet: XSSFSheet, private val row: XSSFRow) : Styleable {
    val rowNumber: Int
        get() = row.rowNum

    var heightInPoints: Float? = null
        set(value) {
            row.heightInPoints = value ?: sheet.defaultRowHeightInPoints
            field = value
        }

    var rowStyle: XSSFCellStyle? = row.rowStyle
        set(value) {
            row.rowStyle = value
            field = value
        }

    private val applyLate = mutableListOf<() -> Unit>()

    fun build() {
        applyLate.forEach { it() }
    }

    fun cell(block: XSSFCellBuilder.() -> Unit) {
        val cell = row.createCell(row.physicalNumberOfCells)
        val builder = XSSFCellBuilder(sheet, row, cell)
        builder.block()
    }

    override fun setStyle(builder: XSSFCellStyleBuilder) {
        row.rowStyle = builder.build(sheet.workbook, row.rowStyle)
    }

    fun setBorderTop(borderStyle: BorderStyle, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderTop(
                borderStyle,
                CellRangeAddress(row.rowNum, row.rowNum, firstCol, lastCol),
                sheet
            )
        }
    }

    fun setBorderBottom(borderStyle: BorderStyle, firstCol: Int, lastCol: Int) {
        applyLate += {
            RegionUtil.setBorderBottom(
                borderStyle,
                CellRangeAddress(row.rowNum, row.rowNum, firstCol, lastCol),
                sheet
            )
        }
    }
}

class XSSFCellBuilder(private val sheet: XSSFSheet, private val row: XSSFRow, private val cell: XSSFCell) : Styleable {
    var cellStyle: XSSFCellStyle? = cell.cellStyle
        set(value) {
            cell.cellStyle = value
            field = value
        }

    fun value(value: String) {
        cell.setCellValue(value)
    }

    operator fun String.unaryPlus() {
        cell.setCellValue(this)
    }

    fun merge(rowSpan: Int, colSpan: Int) {
        if (rowSpan == 1 && colSpan == 1) return

        sheet.addMergedRegionUnsafe(
            CellRangeAddress(
                row.rowNum,
                row.rowNum + rowSpan - 1,
                cell.columnIndex,
                cell.columnIndex + colSpan - 1
            )
        )
    }

    override fun setStyle(builder: XSSFCellStyleBuilder) {
        cell.cellStyle = builder.build(sheet.workbook, cell.cellStyle)
    }
}

interface Styleable {
    fun style(block: XSSFCellStyleBuilder.() -> Unit) {
        val builder = XSSFCellStyleBuilder()
        builder.block()
        setStyle(builder)

    }

    fun setStyle(builder: XSSFCellStyleBuilder)
}

class XSSFCellStyleBuilder() {
    var dataFormat: Short? = null
    var dataFormatString: String? = null
    var font: Font? = null
    var fontIndex: Int? = null
    var hidden: Boolean? = null
    var locked: Boolean? = null
    var quotePrefixed: Boolean? = null
    var alignment: HorizontalAlignment? = null
    var wrapText: Boolean? = null
    var verticalAlignment: VerticalAlignment? = null
    var rotation: Short? = null
    var indention: Short? = null
    var borderLeft: BorderStyle? = null
    var borderRight: BorderStyle? = null
    var borderTop: BorderStyle? = null
    var borderBottom: BorderStyle? = null
    var leftBorderColor: Short? = null
    var rightBorderColor: Short? = null
    var topBorderColor: Short? = null
    var bottomBorderColor: Short? = null
    var fillPattern: FillPatternType? = null
    var fillBackgroundColor: Short? = null
    var fillBackgroundColorColor: Color? = null
    var fillForegroundColor: Short? = null
    var fillForegroundColorColor: Color? = null
    var shrinkToFit: Boolean? = null

    fun build(workbook: XSSFWorkbook, baseStyle: XSSFCellStyle?): XSSFCellStyle {
        val cellStyle = workbook.createCellStyle()

        baseStyle?.let(cellStyle::cloneStyleFrom)

        font?.let(cellStyle::setFont)
        hidden?.let(cellStyle::setHidden)
        locked?.let(cellStyle::setLocked)
        quotePrefixed?.let(cellStyle::setQuotePrefixed)
        alignment?.let(cellStyle::setAlignment)
        wrapText?.let(cellStyle::setWrapText)
        verticalAlignment?.let(cellStyle::setVerticalAlignment)
        rotation?.let(cellStyle::setRotation)
        indention?.let(cellStyle::setIndention)
        borderLeft?.let(cellStyle::setBorderLeft)
        borderRight?.let(cellStyle::setBorderRight)
        borderTop?.let(cellStyle::setBorderTop)
        borderBottom?.let(cellStyle::setBorderBottom)
        leftBorderColor?.let(cellStyle::setLeftBorderColor)
        rightBorderColor?.let(cellStyle::setRightBorderColor)
        topBorderColor?.let(cellStyle::setTopBorderColor)
        bottomBorderColor?.let(cellStyle::setBottomBorderColor)
        fillPattern?.let(cellStyle::setFillPattern)
        fillBackgroundColor?.let(cellStyle::setFillBackgroundColor)
        fillBackgroundColorColor?.let(cellStyle::setFillBackgroundColor)
        fillForegroundColor?.let(cellStyle::setFillForegroundColor)
        fillForegroundColorColor?.let(cellStyle::setFillForegroundColor)

        return cellStyle
    }
}

class XSSFFontBuilder {
    var fontName: String? = null
    var fontHeight: Short? = null
    var fontHeightInPoints: Short? = null
    var bold: Boolean? = null
    var italic: Boolean? = null
    var strikeout: Boolean? = null
    var typeOffset: Short? = null
    var underline: Byte? = null
    var charSet: Byte? = null
    var color: Short? = null

    fun build(workbook: XSSFWorkbook): XSSFFont {
        val font = workbook.createFont()

        fontName?.let(font::setFontName)
        fontHeight?.let(font::setFontHeight)
        fontHeightInPoints?.let(font::setFontHeightInPoints)
        bold?.let(font::setBold)
        italic?.let(font::setItalic)
        strikeout?.let(font::setStrikeout)
        typeOffset?.let(font::setTypeOffset)
        underline?.let(font::setUnderline)
        charSet?.let(font::setCharSet)
        color?.let(font::setColor)

        return font
    }
}