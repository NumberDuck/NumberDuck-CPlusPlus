/*
Number Duck
Copyright (c) 2012-2025 File Scribe

Closed source licenses may be purchased from https://numberduck.com

Otherwise:

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

#pragma once

#include "NumberDuck.hpp"

namespace NumberDuck
{
	namespace Secret
	{
		class ValueImplementation;
		class CellImplementation;
		class StyleImplementation;
		class InternalString;
		class LineImplementation;
		class MarkerImplementation;
		class FillImplementation;
		class FontImplementation;
		class ChartImplementation;
		class Formula;
		class PictureImplementation;
		class WorksheetImplementation;
		class ColumnInfo;
		template <class T>
		class TableElement;
		class RowInfo;
		class Coordinate;
		class JpegLoader;
		class JpegImageInfo;
		class PngLoader;
		class PngImageInfo;
		class OfficeArtRecordHeaderStruct;
		class OfficeArtRecord;
		class Token;
		class WorkbookGlobals;
		class ParsedExpressionRecord;
		class BiffHeader;
		class BiffRecord_ContinueInfo;
		class Stream;
		class BiffRecord;
		class SharedStringContainer;
		class WorksheetRange;
		class BiffWorksheetStreamSize;
		class BiffRecordContainer;
		class MsoDrawingGroupRecord;
		class PaletteRecord;
		class FontRecord;
		class Format;
		class XF;
		class XFExt;
		class BoundSheet8Record;
		class Sector;
		class StreamDataStruct;
		class SectorChain;
		class CompoundFile;
		class OfficeArtFOPTEStruct;
		class MSOCRStruct;
		class OfficeArtFRITStruct;
		class OfficeArtIDCLStruct;
		class MD4DigestStruct;
		class OfficeArtBlipRecord;
		class ShortXLUnicodeStringStruct;
		class RgceLocStruct;
		class Coordinate3d;
		class PtgAttrSpaceTypeStruct;
		class RgceAreaStruct;
		class Area;
		class Area3d;
		class FrtHeaderStruct;
		class ExtPropStruct;
		class Theme;
		class XLUnicodeStringStruct;
		class BuiltInStyleStruct;
		class FrtHeaderOldStruct;
		class XLUnicodeRichExtendedString;
		class RwStruct;
		class FtCmoStruct;
		class FtCfStruct;
		class FtPioGrbitStruct;
		class CellStruct;
		class ColStruct;
		class RkRecStruct;
		class IXFCellStruct;
		class MsoDrawingRecord_Position;
		class OfficeArtDgContainerRecord;
		class OfficeArtSpContainerRecord;
		class OfficeArtDggContainerRecord;
		class Ref8Struct;
		class BiffWorkbookGlobals;
		class OfficeArtFOPTRecord;
		class OfficeArtTertiaryFOPTRecord;
		class FormulaValueStruct;
		class CellParsedFormulaStruct;
		class IcvFontStruct;
		class XTIStruct;
		class XLUnicodeRichExtendedString_ContinueInfo;
		class RwUStruct;
		class ColRelUStruct;
		class OfficeArtFOPTEOPIDStruct;
		class HyperlinkObjectStruct;
		class HyperlinkStringStruct;
		class FrtFlagsStruct;
		class FullColorExtStruct;
		class RgceStruct;
		class RowRecord;
		class OfficeArtDimensions;
		class RedBlackNode;
		class StreamDirectoryImplementation;
		class SectorImplementation;
		class CompoundFileHeader;
		class MasterSectorAllocationTable;
		class SectorAllocationTable;
		class StreamDirectory;
		class ParseFunctionData;
		class ParseSpaceData;
		class RedBlackNodeImplementation;
		class RedBlackTreeImplementation;
		class XlsxWorkbookGlobals;
		class XmlNode;
		class WorkbookImplementation;
		template <class T>
		class OwnedVector;
		class BofRecord;
		class BiffWorksheet;
		class Zip;
		class XmlFile;
		class XlsxWorksheet;
		template <class T>
		class Vector;
		class SharedString;
		class SeriesImplementation;
		class MergedCellImplementation;
		class LegendImplementation;
		class ZipFileInfo;
	}
	class Value;
	class Worksheet;
	class Cell;
	class Style;
	class Font;
	class Color;
	class Line;
	class Marker;
	class Fill;
	class Series;
	class Legend;
	class Blob;
	class Workbook;
	class Picture;
	class Chart;
	class MergedCell;
	class BlobView;
}
namespace NumberDuck
{
	class Value
	{
		public: enum Type
		{
			TYPE_EMPTY,
			TYPE_STRING,
			TYPE_FLOAT,
			TYPE_BOOLEAN,
			TYPE_FORMULA,
			TYPE_AREA,
			TYPE_AREA_3D,
			TYPE_ERROR,
		};

		public: Secret::ValueImplementation* m_pImpl = 0;
		public: Value();
		public: bool Equals(const Value* pValue) const;
		public: Value::Type GetType() const;
		public: const char* GetString() const;
		public: double GetFloat() const;
		public: bool GetBoolean() const;
		public: const char* GetFormula() const;
		public: const Value* EvaulateFormula() const;
		public: virtual ~Value();
	};
}

namespace NumberDuck
{
	class Cell
	{
		public: Secret::CellImplementation* m_pImpl = 0;
		public: Cell(Worksheet* pWorksheet);
		public: bool Equals(const Cell* pCell) const;
		public: const Value* GetValue();
		public: Value::Type GetType();
		public: void Clear();
		public: void SetString(const char* szString);
		public: const char* GetString();
		public: void SetFloat(double fFloat);
		public: double GetFloat();
		public: void SetBoolean(bool bBoolean);
		public: bool GetBoolean();
		public: void SetFormula(const char* szFormula);
		public: const char* GetFormula();
		public: const Value* EvaulateFormula() const;
		public: bool SetStyle(Style* pStyle);
		public: Style* GetStyle();
		public: virtual ~Cell();
	};
}

namespace NumberDuck
{
	class Style
	{
		public: Secret::StyleImplementation* m_pImplementation = 0;
		public: Style();
		public: Font* GetFont();
		public: enum HorizontalAlign
		{
			HORIZONTAL_ALIGN_GENERAL = 0,
			HORIZONTAL_ALIGN_LEFT,
			HORIZONTAL_ALIGN_CENTER,
			HORIZONTAL_ALIGN_RIGHT,
			HORIZONTAL_ALIGN_FILL,
			HORIZONTAL_ALIGN_JUSTIFY,
			HORIZONTAL_ALIGN_CENTER_ACROSS_SELECTION,
			HORIZONTAL_ALIGN_DISTRIBUTED,
		};

		public: Style::HorizontalAlign GetHorizontalAlign();
		public: void SetHorizontalAlign(Style::HorizontalAlign eHorizontalAlign);
		public: enum VerticalAlign
		{
			VERTICAL_ALIGN_TOP = 0,
			VERTICAL_ALIGN_CENTER,
			VERTICAL_ALIGN_BOTTOM,
			VERTICAL_ALIGN_JUSTIFY,
			VERTICAL_ALIGN_DISTRIBUTED,
		};

		public: Style::VerticalAlign GetVerticalAlign();
		public: void SetVerticalAlign(Style::VerticalAlign eVerticalAlign);
		public: Color* GetBackgroundColor(bool bCreateIfMissing);
		public: void ClearBackgroundColor();
		public: enum FillPattern
		{
			FILL_PATTERN_NONE = 0,
			FILL_PATTERN_50,
			FILL_PATTERN_75,
			FILL_PATTERN_25,
			FILL_PATTERN_125,
			FILL_PATTERN_625,
			FILL_PATTERN_HORIZONTAL_STRIPE,
			FILL_PATTERN_VARTICAL_STRIPE,
			FILL_PATTERN_DIAGONAL_STRIPE,
			FILL_PATTERN_REVERSE_DIAGONAL_STRIPE,
			FILL_PATTERN_DIAGONAL_CROSSHATCH,
			FILL_PATTERN_THICK_DIAGONAL_CROSSHATCH,
			FILL_PATTERN_THIN_HORIZONTAL_STRIPE,
			FILL_PATTERN_THIN_VERTICAL_STRIPE,
			FILL_PATTERN_THIN_REVERSE_VERTICAL_STRIPE,
			FILL_PATTERN_THIN_DIAGONAL_STRIPE,
			FILL_PATTERN_THIN_HORIZONTAL_CROSSHATCH,
			FILL_PATTERN_THIN_DIAGONAL_CROSSHATCH,
			NUM_FILL_PATTERN,
		};

		public: Style::FillPattern GetFillPattern();
		public: void SetFillPattern(Style::FillPattern eFillPattern);
		public: Color* GetFillPatternColor(bool bCreateIfMissing);
		public: void ClearFillPatternColor();
		public: Line* GetTopBorderLine();
		public: Line* GetRightBorderLine();
		public: Line* GetBottomBorderLine();
		public: Line* GetLeftBorderLine();
		public: const char* GetFormat();
		public: void SetFormat(const char* szFormat);
		public: virtual ~Style();
	};
}

namespace NumberDuck
{
	class Line
	{
		protected: Secret::LineImplementation* m_pImpl = 0;
		public: enum Type
		{
			TYPE_NONE = 0,
			TYPE_THIN,
			TYPE_DASHED,
			TYPE_DOTTED,
			TYPE_DASH_DOT,
			TYPE_DASH_DOT_DOT,
			TYPE_MEDIUM,
			TYPE_MEDIUM_DASHED,
			TYPE_MEDIUM_DASH_DOT,
			TYPE_MEDIUM_DASH_DOT_DOT,
			TYPE_THICK,
			TYPE_DOUBLE,
			TYPE_HAIR,
			TYPE_SLANT_DASH_DOT_DOT,
			NUM_TYPE,
		};

		public: Line();
		public: bool Equals(const Line* pLine) const;
		public: Line::Type GetType() const;
		public: void SetType(Type eType);
		public: Color* GetColor();
		public: virtual ~Line();
	};
}

namespace NumberDuck
{
	class Marker
	{
		public: Secret::MarkerImplementation* m_pImpl = 0;
		public: enum Type
		{
			TYPE_NONE = 0,
			TYPE_SQUARE,
			TYPE_DIAMOND,
			TYPE_TRIANGLE,
			TYPE_X,
			TYPE_ASTERISK,
			TYPE_SHORT_BAR,
			TYPE_LONG_BAR,
			TYPE_CIRCULAR,
			TYPE_PLUS,
			NUM_TYPE,
		};

		public: Marker();
		public: bool Equals(const Marker* pMarker) const;
		public: Marker::Type GetType() const;
		public: void SetType(Marker::Type eType);
		public: Color* GetFillColor(bool bCreateIfMissing);
		public: void SetFillColor(const Color* pColor);
		public: void ClearFillColor();
		public: Color* GetBorderColor(bool bCreateIfMissing);
		public: void SetBorderColor(const Color* pColor);
		public: void ClearBorderColor();
		public: int GetSize() const;
		public: void SetSize(int nSize);
		public: virtual ~Marker();
	};
}

namespace NumberDuck
{
	class Fill
	{
		protected: Secret::FillImplementation* m_pImpl = 0;
		public: enum Type
		{
			TYPE_NONE = 0,
			TYPE_SOLID,
			NUM_TYPE,
		};

		public: Fill();
		public: bool Equals(const Fill* pFill) const;
		public: Fill::Type GetType() const;
		public: void SetType(Type eType);
		public: Color* GetForegroundColor();
		public: Color* GetBackgroundColor();
		public: virtual ~Fill();
	};
}

namespace NumberDuck
{
	class Font
	{
		public: Secret::FontImplementation* m_pImpl = 0;
		public: enum Underline
		{
			UNDERLINE_NONE = 0,
			UNDERLINE_SINGLE,
			UNDERLINE_DOUBLE,
			UNDERLINE_SINGLE_ACCOUNTING,
			UNDERLINE_DOUBLE_ACCOUNTING,
			NUM_UNDERLINE,
		};

		public: Font();
		public: void SetName(const char* szName);
		public: const char* GetName();
		public: void SetSize(unsigned char nSize);
		public: unsigned char GetSize();
		public: Color* GetColor(bool bCreateIfMissing);
		public: void ClearColor();
		public: bool GetBold();
		public: void SetBold(bool bBold);
		public: bool GetItalic();
		public: void SetItalic(bool bItalic);
		public: Font::Underline GetUnderline();
		public: void SetUnderline(Font::Underline eUnderline);
		public: virtual ~Font();
	};
}

namespace NumberDuck
{
	class Chart
	{
		public: enum Type
		{
			TYPE_COLUMN = 0,
			TYPE_COLUMN_STACKED,
			TYPE_COLUMN_STACKED_100,
			TYPE_BAR,
			TYPE_BAR_STACKED,
			TYPE_BAR_STACKED_100,
			TYPE_LINE,
			TYPE_LINE_STACKED,
			TYPE_LINE_STACKED_100,
			TYPE_AREA,
			TYPE_AREA_STACKED,
			TYPE_AREA_STACKED_100,
			TYPE_SCATTER,
			NUM_TYPE,
		};

		public: Secret::ChartImplementation* m_pImpl = 0;
		public: Chart(Worksheet* pWorksheet, Type eType);
		public: unsigned int GetX();
		public: bool SetX(unsigned int nX);
		public: unsigned int GetY();
		public: bool SetY(unsigned int nY);
		public: unsigned int GetSubX();
		public: void SetSubX(unsigned int nSubX);
		public: unsigned int GetSubY();
		public: void SetSubY(unsigned int nSubY);
		public: unsigned int GetWidth();
		public: void SetWidth(unsigned int nWidth);
		public: unsigned int GetHeight();
		public: void SetHeight(unsigned int nHeight);
		public: Chart::Type GetType();
		public: void SetType(Type eType);
		public: unsigned int GetNumSeries();
		public: Series* GetSeriesByIndex(unsigned int nIndex);
		public: Series* CreateSeries(const char* szValues);
		public: void PurgeSeries(unsigned int nIndex);
		public: const char* GetCategories();
		public: bool SetCategories(const char* szCategories);
		public: const char* GetTitle();
		public: void SetTitle(const char* szTitle);
		public: const char* GetHorizontalAxisLabel();
		public: void SetHorizontalAxisLabel(const char* szHorizontalAxisLabel);
		public: const char* GetVerticalAxisLabel();
		public: void SetVerticalAxisLabel(const char* szVerticalAxisLabel);
		public: Legend* GetLegend();
		public: Line* GetFrameBorderLine();
		public: Fill* GetFrameFill();
		public: Line* GetPlotBorderLine();
		public: Fill* GetPlotFill();
		public: Line* GetHorizontalAxisLine();
		public: Line* GetHorizontalGridLine();
		public: Line* GetVerticalAxisLine();
		public: Line* GetVerticalGridLine();
		public: virtual ~Chart();
	};
}

namespace NumberDuck
{
	class Picture
	{
		public: Secret::PictureImplementation* m_pImplementation = 0;
		public: enum Format
		{
			PNG,
			JPEG,
			EMF,
			WMF,
			PICT,
			DIB,
			TIFF,
		};

		public: Picture(Blob* pBlob, Picture::Format eFormat);
		public: unsigned int GetX();
		public: bool SetX(unsigned int nX);
		public: unsigned int GetY();
		public: bool SetY(unsigned int nY);
		public: unsigned int GetSubX();
		public: void SetSubX(unsigned int nSubX);
		public: unsigned int GetSubY();
		public: void SetSubY(unsigned int nSubY);
		public: unsigned int GetWidth();
		public: void SetWidth(unsigned int nWidth);
		public: unsigned int GetHeight();
		public: void SetHeight(unsigned int nHeight);
		public: const char* GetUrl();
		public: void SetUrl(const char* szUrl);
		public: Blob* GetBlob();
		public: Picture::Format GetFormat();
		public: virtual ~Picture();
	};
}

namespace NumberDuck
{
	class Worksheet
	{
		public: Secret::WorksheetImplementation* m_pImpl = 0;
		public: static const unsigned short MAX_COLUMN = 255;
		public: static const unsigned short MAX_ROW = 65535;
		public: static const unsigned short DEFAULT_COLUMN_WIDTH = 64;
		public: static const unsigned short DEFAULT_ROW_HEIGHT = 20;
		public: enum Orientation
		{
			ORIENTATION_PORTRAIT,
			ORIENTATION_LANDSCAPE,
		};

		public: Worksheet(Workbook* pWorkbook);
		public: virtual ~Worksheet();
		public: const char* GetName();
		public: bool SetName(const char* szName);
		public: Worksheet::Orientation GetOrientation();
		public: void SetOrientation(Worksheet::Orientation eOrientation);
		public: bool GetPrintGridlines();
		public: void SetPrintGridlines(bool bPrintGridlines);
		public: bool GetShowGridlines();
		public: void SetShowGridlines(bool bShowGridlines);
		public: unsigned short GetColumnWidth(unsigned short nColumn);
		public: void SetColumnWidth(unsigned short nColumn, unsigned short nWidth);
		public: bool GetColumnHidden(unsigned short nColumn);
		public: void SetColumnHidden(unsigned short nColumn, bool bHidden);
		public: void InsertColumn(unsigned short nColumn);
		public: void DeleteColumn(unsigned short nColumn);
		public: unsigned short GetRowHeight(unsigned short nRow);
		public: void SetRowHeight(unsigned short nRow, unsigned short nHeight);
		public: void InsertRow(unsigned short nRow);
		public: void DeleteRow(unsigned short nRow);
		public: Cell* GetCell(unsigned short nX, unsigned short nY);
		public: Cell* GetCellByRC(unsigned short nRow, unsigned short nColumn);
		public: Cell* GetCellByAddress(const char* szAddress);
		public: int GetNumPicture();
		public: Picture* GetPictureByIndex(int nIndex);
		public: Picture* CreatePicture(const char* szFileName);
		public: void PurgePicture(int nIndex);
		public: int GetNumChart();
		public: Chart* GetChartByIndex(int nIndex);
		public: Chart* CreateChart(Chart::Type eType);
		public: void PurgeChart(int nIndex);
		public: int GetNumMergedCell();
		public: MergedCell* GetMergedCellByIndex(int nIndex);
		public: MergedCell* CreateMergedCell(unsigned short nX, unsigned short nY, unsigned short nWidth, unsigned short nHeight);
		public: void PurgeMergedCell(int nIndex);
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtRecord
		{
			public: static const unsigned short MAX_DATA_SIZE = 8224;
			public: enum Type
			{
				TYPE_OFFICE_ART_DGG_CONTAINER = 0xF000,
				TYPE_OFFICE_ART_B_STORE_CONTAINER = 0xF001,
				TYPE_OFFICE_ART_DG_CONTAINER = 0xF002,
				TYPE_OFFICE_ART_SPGR_CONTAINER = 0xF003,
				TYPE_OFFICE_ART_SP_CONTAINER = 0xF004,
				TYPE_OFFICE_ART_FDGG_BLOCK = 0xF006,
				TYPE_OFFICE_ART_FBSE = 0xF007,
				TYPE_OFFICE_ART_FDG = 0xF008,
				TYPE_OFFICE_ART_FSPGR = 0xF009,
				TYPE_OFFICE_ART_FSP = 0xF00A,
				TYPE_OFFICE_ART_FOPT = 0xF00B,
				TYPE_OFFICE_ART_TERTIARY_FOPT = 0xF122,
				TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET = 0xF010,
				TYPE_OFFICE_ART_CLIENT_DATA = 0xF011,
				TYPE_OFFICE_ART_BLIP_EMF = 0xF01A,
				TYPE_OFFICE_ART_BLIP_WMF = 0xF01B,
				TYPE_OFFICE_ART_BLIP_PICT = 0xF01C,
				TYPE_OFFICE_ART_BLIP_JPEG = 0xF01D,
				TYPE_OFFICE_ART_BLIP_PNG = 0xF01E,
				TYPE_OFFICE_ART_BLIP_DIB = 0xF01F,
				TYPE_OFFICE_ART_BLIP_TIFF = 0xF029,
				TYPE_OFFICE_ART_BLIP_JPEG_CMYK = 0xF02A,
				TYPE_OFFICE_ART_FRIT_CONTAINER = 0xF118,
				TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER = 0xF11E,
			};

			public: enum OPIDType
			{
				OPID_PROTECTION_BOOLEAN_PROPERTIES = 0x007F,
				OPID_TEXT_BOOLEAN_PROPERTIES = 0x00BF,
				OPID_PIB = 0x0104,
				OPID_PIB_NAME = 0x0105,
				OPID_BLIP_BOOLEAN_PROPERTIES = 0x013F,
				OPID_FILL_COLOR = 0x0181,
				OPID_FILL_OPACITY = 0x0182,
				OPID_FILL_BACK_COLOR = 0x0183,
				OPID_FILL_BACK_OPACITY = 0x0184,
				OPID_FILL_STYLE_BOOLEAN_PROPERTIES = 0x01BF,
				OPID_LINE_COLOR = 0x01C0,
				OPID_LINE_STYLE_BOOLEAN_PROPERTIES = 0x01FF,
				OPID_SHADOW_STYLE_BOOLEAN_PROPERTIES = 0x023F,
				OPID_SHAPE_BOOLEAN_PROPERTIES = 0x033F,
				OPID_WZ_NAME = 0x0380,
				OPID_HYPERLINK = 0x0382,
				OPID_GROUP_SHAPE_BOOLEAN_PROPERTIES = 0x03BF,
			};

			protected: OfficeArtRecordHeaderStruct* m_pHeader = 0;
			protected: bool m_bIsContainer;
			protected: OwnedVector<OfficeArtRecord*>* m_pOfficeArtRecordVector = 0;
			public: const char* GetTypeName();
			public: OfficeArtRecord(OfficeArtRecordHeaderStruct* pHeader, bool bIsContainer, BlobView* pBlobView);
			public: OfficeArtRecord(Type nType, bool bIsContainer, unsigned int nSize, bool bHax);
			public: unsigned short GetVersion();
			public: unsigned short GetInstance();
			public: OfficeArtRecord::Type GetType();
			public: bool GetIsContainer();
			public: unsigned int GetSize();
			public: virtual unsigned int GetRecursiveSize();
			public: virtual void RecursiveWrite(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: unsigned short GetNumOfficeArtRecord();
			public: OfficeArtRecord* GetOfficeArtRecordByIndex(unsigned short nIndex);
			public: OfficeArtRecord* FindOfficeArtRecordByType(Type eType);
			public: virtual void AddOfficeArtRecord(OfficeArtRecord* pOfficeArtRecord);
			public: static OfficeArtRecord* CreateOfficeArtRecord(BlobView* pBlobView);
			public: virtual ~OfficeArtRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ParsedExpressionRecord
		{
			public: enum Type
			{
				TYPE_PtgExp,
				TYPE_PtgAdd,
				TYPE_PtgSub,
				TYPE_PtgMul,
				TYPE_PtgDiv,
				TYPE_PtgPower,
				TYPE_PtgConcat,
				TYPE_PtgLt,
				TYPE_PtgLe,
				TYPE_PtgEq,
				TYPE_PtgGe,
				TYPE_PtgGt,
				TYPE_PtgNe,
				TYPE_PtgParen,
				TYPE_PtgMissArg,
				TYPE_PtgStr,
				TYPE_PtgAttrSemi,
				TYPE_PtgAttrIf,
				TYPE_PtgAttrGoto,
				TYPE_PtgAttrSum,
				TYPE_PtgAttrSpace,
				TYPE_PtgBool,
				TYPE_PtgInt,
				TYPE_PtgNum,
				TYPE_PtgArea,
				TYPE_PtgArea3d,
				TYPE_PtgFunc,
				TYPE_PtgFuncVar,
				TYPE_PtgRef,
				TYPE_PtgRef3d,
				TYPE_UNKNOWN,
			};

			protected: Type m_eType;
			protected: unsigned short m_nSize;
			protected: unsigned char m_nFirstByte;
			protected: unsigned char m_nSecondByte;
			public: ParsedExpressionRecord(Type eType, unsigned short nSize, bool bHax);
			public: virtual ~ParsedExpressionRecord();
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: ParsedExpressionRecord::Type GetType();
			public: const char* GetTypeName();
			public: unsigned short GetDataSize();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static ParsedExpressionRecord* CreateParsedExpressionRecord(BlobView* pBlobView);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffHeader
		{
			public: unsigned short m_nType;
			public: unsigned int m_nSize;
		};
		class BiffRecord
		{
			public: static const unsigned short MAX_DATA_SIZE = 8224;
			public: enum Type
			{
				TYPE_DIMENSIONS_B2 = 0x0000,
				TYPE_BLANK_B2 = 0x0001,
				TYPE_INTEGER_B2_ONLY = 0x0002,
				TYPE_NUMBER_B2 = 0x0003,
				TYPE_LABEL_B2 = 0x0004,
				TYPE_BOOLERR_B2 = 0x0005,
				TYPE_FORMULA = 0x0006,
				TYPE_STRING_B2 = 0x0007,
				TYPE_ROW_B2 = 0x0008,
				TYPE_BOF_B2 = 0x0009,
				TYPE_EOF = 0x000A,
				TYPE_INDEX_B2_ONLY = 0x000B,
				TYPE_CALCCOUNT = 0x000C,
				TYPE_CALCMODE = 0x000D,
				TYPE_CALC_PRECISION = 0x000E,
				TYPE_REFMODE = 0x000F,
				TYPE_DELTA = 0x0010,
				TYPE_ITERATION = 0x0011,
				TYPE_PROTECT = 0x0012,
				TYPE_PASSWORD = 0x0013,
				TYPE_HEADER = 0x0014,
				TYPE_FOOTER = 0x0015,
				TYPE_EXTERNCOUNT = 0x0016,
				TYPE_EXTERN_SHEET = 0x0017,
				TYPE_NAME_B2 = 0x0018,
				TYPE_WIN_PROTECT = 0x0019,
				TYPE_VERTICALPAGEBREAKS = 0x001A,
				TYPE_HORIZONTALPAGEBREAKS = 0x001B,
				TYPE_NOTE = 0x001C,
				TYPE_SELECTION = 0x001D,
				TYPE_FORMAT_B2 = 0x001E,
				TYPE_BUILTINFMTCOUNT_B2 = 0x001F,
				TYPE_COLUMNDEFAULT_B2_ONLY = 0x0020,
				TYPE_ARRAY_B2_ONLY = 0x0021,
				TYPE_DATE_1904 = 0x0022,
				TYPE_EXTERN_NAME = 0x0023,
				TYPE_COLWIDTH_B2_ONLY = 0x0024,
				TYPE_DEFAULTROWHEIGHT_B2_ONLY = 0x0025,
				TYPE_LEFT_MARGIN = 0x0026,
				TYPE_RIGHT_MARGIN = 0x0027,
				TYPE_TOP_MARGIN = 0x0028,
				TYPE_BOTTOM_MARGIN = 0x0029,
				TYPE_PrintRowCol = 0x002A,
				TYPE_PrintGrid = 0x002B,
				TYPE_FILEPASS = 0x002F,
				TYPE_FONT = 0x0031,
				TYPE_FONT2_B2_ONLY = 0x0032,
				TYPE_PRINT_SIZE = 0x0033,
				TYPE_TABLEOP_B2 = 0x0036,
				TYPE_TABLEOP2_B2 = 0x0037,
				TYPE_CONTINUE = 0x003C,
				TYPE_WINDOW1 = 0x003D,
				TYPE_WINDOW2_B2 = 0x003E,
				TYPE_BACKUP = 0x0040,
				TYPE_PANE = 0x0041,
				TYPE_CODE_PAGE = 0x0042,
				TYPE_XF_B2 = 0x0043,
				TYPE_IXFE_B2_ONLY = 0x0044,
				TYPE_EFONT_B2_ONLY = 0x0045,
				TYPE_PLS = 0x004D,
				TYPE_DCONREF = 0x0051,
				TYPE_DEFCOLWIDTH = 0x0055,
				TYPE_BUILTINFMTCOUNT_B3 = 0x0056,
				TYPE_XCT = 0x0059,
				TYPE_CRN = 0x005A,
				TYPE_FILESHARING = 0x005B,
				TYPE_WRITE_ACCESS = 0x005C,
				TYPE_OBJ = 0x005D,
				TYPE_UNCALCED = 0x005E,
				TYPE_SAVERECALC = 0x005F,
				TYPE_OBJECTPROTECT = 0x0063,
				TYPE_COLINFO = 0x007D,
				TYPE_RK2_mythical = 0x007E,
				TYPE_GUTS = 0x0080,
				TYPE_WSBOOL = 0x0081,
				TYPE_GRIDSET = 0x0082,
				TYPE_HCENTER = 0x0083,
				TYPE_VCENTER = 0x0084,
				TYPE_BOUND_SHEET_8 = 0x0085,
				TYPE_WRITEPROT = 0x0086,
				TYPE_COUNTRY = 0x008C,
				TYPE_HIDE_OBJ = 0x008D,
				TYPE_SHEETSOFFSET = 0x008E,
				TYPE_SHEETHDR = 0x008F,
				TYPE_SORT = 0x0090,
				TYPE_PALETTE = 0x0092,
				TYPE_STANDARDWIDTH = 0x0099,
				TYPE_FILTERMODE = 0x009B,
				TYPE_BUILT_IN_FN_GROUP_COUNT = 0x009C,
				TYPE_AUTOFILTERINFO = 0x009D,
				TYPE_AUTOFILTER = 0x009E,
				TYPE_SCL = 0x00A0,
				TYPE_SETUP = 0x00A1,
				TYPE_GCW = 0x00AB,
				TYPE_MULRK = 0x00BD,
				TYPE_MULBLANK = 0x00BE,
				TYPE_MMS = 0x00C1,
				TYPE_RSTRING = 0x00D6,
				TYPE_DBCELL = 0x00D7,
				TYPE_BOOK_BOOL = 0x00DA,
				TYPE_SCENPROTECT = 0x00DD,
				TYPE_XF = 0x00E0,
				TYPE_INTERFACE_HDR = 0x00E1,
				TYPE_INTERFACE_END = 0x00E2,
				TYPE_MergeCells = 0x00E5,
				TYPE_BITMAP = 0x00E9,
				TYPE_MSO_DRAWING_GROUP = 0x00EB,
				TYPE_MSO_DRAWING = 0x00EC,
				TYPE_MSO_DRAWING_SELECTION = 0x00ED,
				TYPE_PHONETIC = 0x00EF,
				TYPE_SST = 0x00FC,
				TYPE_LABELSST = 0x00FD,
				TYPE_EXT_SST = 0x00FF,
				TYPE_RR_TAB_ID = 0x013D,
				TYPE_LABELRANGES = 0x015F,
				TYPE_USESELFS = 0x0160,
				TYPE_DSF = 0x0161,
				TYPE_SUP_BOOK = 0x01AE,
				TYPE_PROT_4_REV = 0x01AF,
				TYPE_CONDFMT = 0x01B0,
				TYPE_CF = 0x01B1,
				TYPE_DVAL = 0x01B2,
				TYPE_TXO = 0x01B6,
				TYPE_REFRESH_ALL = 0x01B7,
				TYPE_HLINK = 0x01B8,
				TYPE_PROT_4_REV_PASS = 0x01BC,
				TYPE_DV = 0x01BE,
				TYPE_EXCEL9_FILE = 0x01C0,
				TYPE_RECALCID = 0x01C1,
				TYPE_DIMENSION = 0x0200,
				TYPE_BLANK = 0x0201,
				TYPE_NUMBER = 0x0203,
				TYPE_LABEL = 0x0204,
				TYPE_BOOLERR = 0x0205,
				TYPE_FORMULA_B3 = 0x0206,
				TYPE_STRING = 0x0207,
				TYPE_ROW = 0x0208,
				TYPE_INDEX_B3 = 0x020B,
				TYPE_NAME = 0x0218,
				TYPE_ARRAY = 0x0221,
				TYPE_EXTERNNAME_B3 = 0x0223,
				TYPE_DEFAULTROWHEIGHT = 0x0225,
				TYPE_FONT_B3B4 = 0x0231,
				TYPE_TABLEOP = 0x0236,
				TYPE_WINDOW2 = 0x023E,
				TYPE_XF_B3 = 0x0243,
				TYPE_RK = 0x027E,
				TYPE_STYLE = 0x0293,
				TYPE_STYLE_EXT = 0x0892,
				TYPE_FORMULA_B4 = 0x0406,
				TYPE_FORMAT = 0x041E,
				TYPE_XF_B4 = 0x0443,
				TYPE_SHRFMLA = 0x04BC,
				TYPE_QUICKTIP = 0x0800,
				TYPE_BOF = 0x0809,
				TYPE_SHEETLAYOUT = 0x0862,
				TYPE_SHEETPROTECTION = 0x0867,
				TYPE_RANGEPROTECTION = 0x0868,
				TYPE_XF_CRC = 0x087C,
				TYPE_XF_EXT = 0x087D,
				TYPE_BOOK_EXT = 0x0863,
				TYPE_THEME = 0x0896,
				TYPE_COMPRESS_PICTURES = 0x089B,
				TYPE_PLV = 0x088B,
				TYPE_COMPAT12 = 0x088C,
				TYPE_MTR_SETTINGS = 0x089A,
				TYPE_FORCE_FULL_CALCULATION = 0x08A3,
				TYPE_TABLE_STYLES = 0x088E,
				TYPE_ChartColors = 684,
				TYPE_ChartFrtInfo = 2128,
				TYPE_StartBlock = 2130,
				TYPE_EndBlock = 2131,
				TYPE_StartObject = 2132,
				TYPE_EndObject = 2133,
				TYPE_CatLab = 2134,
				TYPE_YMult = 2135,
				TYPE_FrtFontList = 2138,
				TYPE_DataLabExt = 2154,
				TYPE_DataLabExtContents = 2155,
				TYPE_NamePublish = 2195,
				TYPE_HeaderFooter = 2204,
				TYPE_CrtMlFrt = 2206,
				TYPE_ShapePropsStream = 2212,
				TYPE_TextPropsStream = 2213,
				TYPE_CrtLayout12A = 2215,
				TYPE_Units = 4097,
				TYPE_Chart = 4098,
				TYPE_Series = 4099,
				TYPE_DataFormat = 4102,
				TYPE_LineFormat = 4103,
				TYPE_MarkerFormat = 4105,
				TYPE_AreaFormat = 4106,
				TYPE_PieFormat = 4107,
				TYPE_AttachedLabel = 4108,
				TYPE_SeriesText = 4109,
				TYPE_ChartFormat = 4116,
				TYPE_Legend = 4117,
				TYPE_SeriesList = 4118,
				TYPE_Bar = 4119,
				TYPE_Line = 4120,
				TYPE_Pie = 4121,
				TYPE_Area = 4122,
				TYPE_Scatter = 4123,
				TYPE_CrtLine = 4124,
				TYPE_Axis = 4125,
				TYPE_Tick = 4126,
				TYPE_ValueRange = 4127,
				TYPE_CatSerRange = 4128,
				TYPE_AxisLine = 4129,
				TYPE_CrtLink = 4130,
				TYPE_DefaultText = 4132,
				TYPE_Text = 4133,
				TYPE_FontX = 4134,
				TYPE_ObjectLink = 4135,
				TYPE_Frame = 4146,
				TYPE_Begin = 4147,
				TYPE_End = 4148,
				TYPE_PlotArea = 4149,
				TYPE_Chart3d = 4154,
				TYPE_PicF = 4156,
				TYPE_DropBar = 4157,
				TYPE_Radar = 4158,
				TYPE_Surf = 4159,
				TYPE_RadarArea = 4160,
				TYPE_AxisParent = 4161,
				TYPE_LegendException = 4163,
				TYPE_ShtProps = 4164,
				TYPE_SerToCrt = 4165,
				TYPE_AxesUsed = 4166,
				TYPE_SerParent = 4170,
				TYPE_SerAuxTrend = 4171,
				TYPE_IFmtRecord = 4174,
				TYPE_Pos = 4175,
				TYPE_AlRuns = 4176,
				TYPE_BRAI = 4177,
				TYPE_BOFDatasheet = 4178,
				TYPE_ExcludeRows = 4179,
				TYPE_ExcludeColumns = 4180,
				TYPE_Orient = 4181,
				TYPE_WinDoc = 4183,
				TYPE_MaxStatus = 4184,
				TYPE_MainWindow = 4185,
				TYPE_SerAuxErrBar = 4187,
				TYPE_ClrtClient = 4188,
				TYPE_SerFmt = 4189,
				TYPE_LinkedSelection = 4190,
				TYPE_Chart3DBarShape = 4191,
				TYPE_Fbi = 4192,
				TYPE_BopPop = 4193,
				TYPE_AxcExt = 4194,
				TYPE_Dat = 4195,
				TYPE_PlotGrowth = 4196,
				TYPE_SIIndex = 4197,
				TYPE_GelFrame = 4198,
				TYPE_BopPopCustom = 4199,
				TYPE_Fbi2 = 4200,
			};

			public: const char* GetTypeName();
			public: static const int SIZEOF_HEADER = 2 + 2;
			protected: BiffHeader* m_pHeader = 0;
			protected: Blob* m_pContinueBlob = 0;
			protected: BlobView* m_pBlobView = 0;
			protected: OwnedVector<BiffRecord_ContinueInfo*>* m_pContinueInfoVector = 0;
			public: BiffRecord(BiffHeader* pHeader, Stream* pStream);
			public: BiffRecord(Type nType, unsigned int nSize);
			public: BiffRecord::Type GetType();
			public: unsigned int GetSize();
			public: static BiffRecord* CreateBiffRecord(Stream* pStream);
			public: virtual void Write(Stream* pStream, BlobView* pTempBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: void Extend(BiffRecord* pBiffRecord);
			public: virtual ~BiffRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffWorksheetStreamSize
		{
			public: unsigned int m_nSize;
		};
		class WorkbookGlobals
		{
			public: SharedStringContainer* m_pSharedStringContainer = 0;
			public: Vector<Picture*>* m_pSharedPictureVector = 0;
			public: OwnedVector<WorksheetRange*>* m_pWorksheetRangeVector = 0;
			public: OwnedVector<BiffWorksheetStreamSize*>* m_pBiffWorksheetStreamSizeVector = 0;
			public: OwnedVector<Style*>* m_pStyleVector = 0;
			public: Font* m_pHeaderFont = 0;
			public: unsigned short m_Window1_nTabRatio;
			public: WorkbookGlobals();
			public: virtual ~WorkbookGlobals();
			public: void Clear();
			public: void PushBiffWorksheetStreamSize(unsigned int nStreamSize);
			public: const char* GetSharedStringByIndex(unsigned int nIndex);
			public: unsigned int GetSharedStringIndex(const char* szString);
			public: unsigned int PushSharedString(const char* szString);
			public: unsigned int PushPicture(Picture* pPicture);
			public: WorksheetRange* GetWorksheetRangeByIndex(unsigned short nIndex);
			public: unsigned short GetWorksheetRangeIndex(unsigned short nFirst, unsigned short nLast);
			public: Style* GetStyleByIndex(unsigned short nIndex);
			public: unsigned short GetStyleIndex(Style* pStyle);
			public: unsigned short GetNumStyle();
			public: Style* CreateStyle();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffWorkbookGlobals : public WorkbookGlobals
		{
			public: static const unsigned short NUM_DEFAULT_PALETTE_ENTRY = 8;
			public: static const unsigned short NUM_CUSTOM_PALETTE_ENTRY = 56;
			public: static const unsigned short PALETTE_INDEX_ERROR = 0xFF;
			public: static const unsigned short PALETTE_INDEX_DEFAULT = 0xFE;
			public: static const unsigned char PALETTE_INDEX_DEFAULT_FOREGROUND = 0x0040;
			public: static const unsigned char PALETTE_INDEX_DEFAULT_BACKGROUND = 0x0041;
			public: static const unsigned short PALETTE_INDEX_DEFAULT_CHART_FOREGROUND = 0x004D;
			public: static const unsigned short PALETTE_INDEX_DEFAULT_CHART_BACKGROUND = 0x004E;
			public: static const unsigned short PALETTE_INDEX_DEFAULT_NEUTRAL = 0x004F;
			public: static const unsigned short PALETTE_INDEX_DEFAULT_TOOL_TIP_TEXT = 0x0051;
			public: static const unsigned short PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC = 0x7FFF;
			public: static const unsigned int DEFAULT_COLOR[NUM_DEFAULT_PALETTE_ENTRY];
			public: static const unsigned int DEFAULT_CUSTOM_COLOR[NUM_CUSTOM_PALETTE_ENTRY];
			protected: BiffRecordContainer* m_pBiffRecordContainer = 0;
			protected: OwnedVector<InternalString*>* m_sWorkheetNameVector = 0;
			protected: InternalString* m_sTempWorksheetName = 0;
			protected: MsoDrawingGroupRecord* m_MsoDrawingGroupRecord = 0;
			protected: PaletteRecord* m_pPaletteRecord = 0;
			public: BiffWorkbookGlobals(BiffRecord* pInitialBiffRecord, Stream* pStream);
			public: static void Write(WorkbookGlobals* pWorkbookGlobals, OwnedVector<Worksheet*>* pWorksheetVector, Stream* pStream);
			public: const char* GetNextWorksheetName();
			public: Style* GetStyleByXfIndex(unsigned short nXfIndex);
			public: MsoDrawingGroupRecord* GetMsoDrawingGroupRecord();
			public: static unsigned char SnapToPalette(Color* pColor);
			public: static unsigned int GetDefaultPaletteColorByIndex(unsigned short nIndex);
			public: unsigned int GetPaletteColorByIndex(unsigned short nIndex);
			public: virtual ~BiffWorkbookGlobals();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffStruct
		{
			public: enum FontScheme
			{
				XFSNONE = 0x00,
				XFSMAJOR = 0x01,
				XFSMINOR = 0x02,
				XFSNIL = 0xFF,
			};

			public: enum XColorType
			{
				XCLRAUTO = 0x00000000,
				XCLRINDEXED = 0x00000001,
				XCLRRGB = 0x00000002,
				XCLRTHEMED = 0x00000003,
				XCLRNINCHED = 0x00000004,
			};

			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtCmoStruct : public BiffStruct
		{
			public: FtCmoStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 22;
			public: void SetDefaults();
			public: unsigned short m_ft;
			public: unsigned short m_cb;
			public: unsigned short m_ot;
			public: unsigned short m_id;
			public: unsigned short m_fLocked;
			public: unsigned short m_reserved;
			public: unsigned short m_fDefaultSize;
			public: unsigned short m_fPublished;
			public: unsigned short m_fPrint;
			public: unsigned short m_unused1;
			public: unsigned short m_unused2;
			public: unsigned short m_fDisabled;
			public: unsigned short m_fUIObj;
			public: unsigned short m_fRecalcObj;
			public: unsigned short m_unused3;
			public: unsigned short m_unused4;
			public: unsigned short m_fRecalcObjAlways;
			public: unsigned short m_unused5;
			public: unsigned short m_unused6;
			public: unsigned short m_unused7;
			public: unsigned int m_unused8;
			public: unsigned int m_unused9;
			public: unsigned int m_unused10;
			public: enum ObjType
			{
				OBJ_TYPE_GROUP = 0x0000,
				OBJ_TYPE_LINE = 0x0001,
				OBJ_TYPE_RECTANGLE = 0x0002,
				OBJ_TYPE_OVAL = 0x0003,
				OBJ_TYPE_ARC = 0x0004,
				OBJ_TYPE_CHART = 0x0005,
				OBJ_TYPE_TEXT = 0x0006,
				OBJ_TYPE_BUTTON = 0x0007,
				OBJ_TYPE_PICTURE = 0x0008,
				OBJ_TYPE_POLYGON = 0x0009,
				OBJ_TYPE_CHECKBOX = 0x000B,
				OBJ_TYPE_RADIO_BUTTON = 0x000C,
				OBJ_TYPE_EDIT_BOX = 0x000D,
				OBJ_TYPE_LABEL = 0x000E,
				OBJ_TYPE_DIALOG_BOX = 0x000F,
				OBJ_TYPE_SPIN_CONTROL = 0x0010,
				OBJ_TYPE_SCROLLBAR = 0x0011,
				OBJ_TYPE_LIST = 0x0012,
				OBJ_TYPE_GROUP_BOX = 0x0013,
				OBJ_TYPE_DROPDOWN_LIST = 0x0014,
				OBJ_TYPE_NOTE = 0x0019,
				OBJ_TYPE_OFFICE_ART_OBJECT = 0x001E,
			};

		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorChain
		{
			protected: int m_nSectorSize;
			protected: Vector<Sector*>* m_pSectorVector = 0;
			protected: Blob* m_pBlob = 0;
			public: SectorChain(int nSectorSize);
			public: int GetNumSector();
			public: Sector* GetSectorByIndex(int nIndex);
			public: int GetSectorSize();
			public: int GetDataSize();
			public: void Read(int nSize, BlobView* pBlobView);
			public: void ReadAt(int nOffset, int nSize, BlobView* pBlobView);
			public: void Write(int nSize, BlobView* pBlobView);
			public: void WriteAt(int nOffset, int nSize, BlobView* pBlobView);
			public: int GetOffset();
			public: void SetOffset(int nOffset);
			public: virtual void AppendSector(Sector* pSector);
			public: virtual void Extend(Sector* pSector);
			public: BlobView* GetBlobView();
			public: void WriteToSectors();
			public: virtual ~SectorChain();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDataStruct
		{
			public: static const int MAX_NAME_LENGTH = 32;
			public: unsigned short m_pName[MAX_NAME_LENGTH];
			public: unsigned short m_nNameDataSize;
			public: unsigned char m_nType;
			public: unsigned char m_nNodeColour;
			public: int m_nLeftChildNodeStreamId;
			public: int m_nRightChildNodeStreamId;
			public: int m_nRootNodeStreamId;
			public: unsigned char m_pUniqueIdentifier[16];
			public: unsigned char m_pUserFlags[4];
			public: unsigned char m_pCreationDate[8];
			public: unsigned char m_pModificationDate[8];
			public: int m_nSectorId;
			public: unsigned int m_nStreamSize;
			public: unsigned char m_pUnused[4];
		};
		class Stream
		{
			public: enum Type
			{
				TYPE_EMPTY = 0,
				TYPE_USER_STORAGE,
				TYPE_USER_STREAM,
				TYPE_LOCK_BYTES,
				TYPE_PROPERTY,
				TYPE_ROOT_STORAGE,
			};

			public: static const int DATA_SIZE = 128;
			protected: int m_nStreamId;
			protected: int m_nMinimumStandardStreamSize;
			protected: StreamDataStruct* m_pDataStruct = 0;
			protected: BlobView* m_pBlobView = 0;
			protected: SectorChain* m_pSectorChain = 0;
			protected: InternalString* m_sNameTemp = 0;
			protected: CompoundFile* m_pCompoundFile = 0;
			public: Stream(int nStreamId, int nMinimumStandardStreamSize, Blob* pBlob, int nOffset, CompoundFile* pCompoundFile);
			public: void Allocate(Type eType, int nStreamSize);
			public: void FillSectorChain();
			public: void WriteToSectors();
			public: int GetStreamId();
			public: const char* GetName();
			public: void SetName(const char* sxName);
			public: Stream::Type GetType();
			public: int GetSectorId();
			public: bool GetShortSector();
			public: int GetStreamSize();
			public: int GetSize();
			public: SectorChain* GetSectorChain();
			public: int GetOffset();
			public: void SetOffset(int nOffset);
			public: void SizeToFit(int nSize);
			public: bool Resize(int nSize);
			public: void SetNodeColour(unsigned char nColour);
			public: void SetLeftChildNodeStreamId(int nStreamId);
			public: void SetRightChildNodeStreamId(int nStreamId);
			public: void SetRootNodeStreamId(int nStreamId);
			public: unsigned short GetNameLengthUtf16();
			public: unsigned short GetNameUtf16(unsigned short nIndex);
			public: virtual ~Stream();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtTertiaryFOPTRecord : public OfficeArtRecord
		{
			public: OfficeArtTertiaryFOPTRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_TERTIARY_FOPT;
			protected: void SetDefaults();
			protected: OwnedVector<OfficeArtFOPTEStruct*>* m_pFoptVector = 0;
			public: OfficeArtTertiaryFOPTRecord();
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: void AddProperty(unsigned short opid, unsigned char fBid, int op);
			public: OfficeArtFOPTEStruct* GetProperty(OfficeArtRecord::OPIDType eType);
			public: virtual ~OfficeArtTertiaryFOPTRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSplitMenuColorContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtSplitMenuColorContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER;
			protected: void SetDefaults();
			protected: MSOCRStruct* m_smca0 = 0;
			protected: MSOCRStruct* m_smca1 = 0;
			protected: MSOCRStruct* m_smca2 = 0;
			protected: MSOCRStruct* m_smca3 = 0;
			public: OfficeArtSplitMenuColorContainerRecord();
			public: virtual ~OfficeArtSplitMenuColorContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSpgrContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtSpgrContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = true;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_SPGR_CONTAINER;
			protected: void SetDefaults();
			public: OfficeArtSpgrContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSpContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtSpContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = true;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_SP_CONTAINER;
			protected: void SetDefaults();
			public: OfficeArtSpContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFSPRecord : public OfficeArtRecord
		{
			public: OfficeArtFSPRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FSP;
			protected: void SetDefaults();
			protected: unsigned int m_spid;
			protected: unsigned int m_fGroup;
			protected: unsigned int m_fChild;
			protected: unsigned int m_fPatriarch;
			protected: unsigned int m_fDeleted;
			protected: unsigned int m_fOleShape;
			protected: unsigned int m_fHaveMaster;
			protected: unsigned int m_fFlipH;
			protected: unsigned int m_fFlipV;
			protected: unsigned int m_fConnector;
			protected: unsigned int m_fHaveAnchor;
			protected: unsigned int m_fBackground;
			protected: unsigned int m_fHaveSpt;
			protected: unsigned int m_unused1;
			public: OfficeArtFSPRecord(unsigned short recInstance, unsigned int spid, unsigned char fGroup, unsigned char fPatriarch, unsigned char fHaveAnchor, unsigned char fHaveSpt);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFSPGRRecord : public OfficeArtRecord
		{
			public: OfficeArtFSPGRRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FSPGR;
			protected: void SetDefaults();
			protected: unsigned int m_xLeft;
			protected: unsigned int m_yTop;
			protected: unsigned int m_xRight;
			protected: unsigned int m_yBottom;
			public: OfficeArtFSPGRRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFRITContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtFRITContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FRIT_CONTAINER;
			protected: void SetDefaults();
			protected: OwnedVector<OfficeArtFRITStruct*>* m_pRgfritVector = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: unsigned short GetNumFRIT();
			public: OfficeArtFRITStruct* GetFRITByIndex(unsigned short nIndex);
			public: virtual ~OfficeArtFRITContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTRecord : public OfficeArtRecord
		{
			public: OfficeArtFOPTRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FOPT;
			protected: void SetDefaults();
			protected: OwnedVector<OfficeArtFOPTEStruct*>* m_pFoptVector = 0;
			public: OfficeArtFOPTRecord();
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: void AddProperty(unsigned short opid, unsigned char fBid, int op);
			public: void AddStringProperty(unsigned short opid, const char* szString);
			public: void AddBlobProperty(unsigned short opid, unsigned char fBid, Blob* pBlob);
			public: OfficeArtFOPTEStruct* GetProperty(OfficeArtRecord::OPIDType eType);
			public: virtual ~OfficeArtFOPTRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFDGRecord : public OfficeArtRecord
		{
			public: OfficeArtFDGRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FDG;
			protected: void SetDefaults();
			protected: unsigned int m_csp;
			protected: unsigned int m_spidCur;
			public: OfficeArtFDGRecord(unsigned short recInstance, unsigned int csp, unsigned int spidCur);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFDGGBlockRecord : public OfficeArtRecord
		{
			public: OfficeArtFDGGBlockRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FDGG_BLOCK;
			protected: void SetDefaults();
			protected: unsigned int m_spidMax;
			protected: unsigned int m_cidcl;
			protected: unsigned int m_cspSaved;
			protected: unsigned int m_cdgSaved;
			protected: OwnedVector<OfficeArtIDCLStruct*>* m_pRgidclVector = 0;
			public: OfficeArtFDGGBlockRecord(unsigned int nNumPicture);
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: virtual ~OfficeArtFDGGBlockRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFBSERecord : public OfficeArtRecord
		{
			public: OfficeArtFBSERecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 36;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_FBSE;
			protected: void SetDefaults();
			protected: unsigned char m_btWin32;
			protected: unsigned char m_btMacOS;
			protected: MD4DigestStruct* m_rgbUid = 0;
			protected: unsigned short m_tag;
			protected: unsigned int m_size;
			protected: unsigned int m_cRef;
			protected: unsigned int m_foDelay;
			protected: unsigned char m_unused1;
			protected: unsigned char m_cbName;
			protected: unsigned char m_unused2;
			protected: unsigned char m_unused3;
			protected: OfficeArtBlipRecord* m_pEmbeddedBlip = 0;
			public: OfficeArtFBSERecord(Picture* pPicture);
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: OfficeArtBlipRecord* GetEmbeddedBlip();
			public: virtual ~OfficeArtFBSERecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDggContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtDggContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = true;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_DGG_CONTAINER;
			protected: void SetDefaults();
			public: OfficeArtDggContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDgContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtDgContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = true;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_DG_CONTAINER;
			protected: void SetDefaults();
			public: OfficeArtDgContainerRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtClientDataRecord : public OfficeArtRecord
		{
			public: OfficeArtClientDataRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_CLIENT_DATA;
			protected: void SetDefaults();
			public: OfficeArtClientDataRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtClientAnchorSheetRecord : public OfficeArtRecord
		{
			public: OfficeArtClientAnchorSheetRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET;
			protected: void SetDefaults();
			protected: unsigned short m_fMove;
			protected: unsigned short m_fSize;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_reserved3;
			protected: unsigned short m_unused;
			protected: unsigned short m_colL;
			protected: unsigned short m_dxL;
			protected: unsigned short m_rwT;
			protected: unsigned short m_dyT;
			protected: unsigned short m_colR;
			protected: unsigned short m_dxR;
			protected: unsigned short m_rwB;
			protected: unsigned short m_dyB;
			public: OfficeArtClientAnchorSheetRecord(unsigned short colL, unsigned short dxL, unsigned short rwT, unsigned short dyT, unsigned short colR, unsigned short dxR, unsigned short rwB, unsigned short dyB);
			public: unsigned short GetCellX1();
			public: unsigned short GetSubCellX1();
			public: unsigned short GetCellY1();
			public: unsigned short GetSubCellY1();
			public: unsigned short GetCellX2();
			public: unsigned short GetSubCellX2();
			public: unsigned short GetCellY2();
			public: unsigned short GetSubCellY2();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtBlipRecord : public OfficeArtRecord
		{
			public: OfficeArtBlipRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const bool IS_CONTAINER = false;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_BLIP_EMF;
			protected: void SetDefaults();
			protected: MD4DigestStruct* m_rgbUid1 = 0;
			protected: MD4DigestStruct* m_rgbUid2 = 0;
			protected: unsigned char m_tag;
			protected: Blob* m_BLIPFileData_ = 0;
			protected: Blob* m_OwnedBLIPFileData = 0;
			public: OfficeArtBlipRecord(Picture* pPicture);
			protected: void PostSetDefaults();
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: Blob* GetBlob();
			public: virtual ~OfficeArtBlipRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtBStoreContainerRecord : public OfficeArtRecord
		{
			public: OfficeArtBStoreContainerRecord(OfficeArtRecordHeaderStruct* pHeader, BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const bool IS_CONTAINER = true;
			protected: static const OfficeArtRecord::Type TYPE = OfficeArtRecord::Type::TYPE_OFFICE_ART_B_STORE_CONTAINER;
			protected: void SetDefaults();
			public: OfficeArtBStoreContainerRecord();
			protected: void PreBlobWrite(BlobView* pBlobView);
			public: virtual void AddOfficeArtRecord(OfficeArtRecord* pOfficeArtRecord);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgStrRecord : public ParsedExpressionRecord
		{
			public: PtgStrRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 1;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgStr;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: ShortXLUnicodeStringStruct* m_string = 0;
			public: PtgStrRecord(const char* szString);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgStrRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgRefRecord : public ParsedExpressionRecord
		{
			public: PtgRefRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 5;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgRef;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved;
			protected: RgceLocStruct* m_loc = 0;
			public: PtgRefRecord(Coordinate* pCoordinate);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgRefRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgRef3dRecord : public ParsedExpressionRecord
		{
			public: PtgRef3dRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 7;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgRef3d;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved;
			protected: unsigned short m_ixti;
			protected: RgceLocStruct* m_loc = 0;
			public: PtgRef3dRecord(Coordinate3d* pCoordinate3d, WorkbookGlobals* pWorkbookGlobals);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgRef3dRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgParenRecord : public ParsedExpressionRecord
		{
			public: PtgParenRecord();
			public: PtgParenRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 1;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgParen;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			public: PtgParenRecord(unsigned char ptg);
			public: void PostBlobRead(BlobView* pBlobView);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgOperatorRecord : public ParsedExpressionRecord
		{
			public: PtgOperatorRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 1;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_UNKNOWN;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			public: PtgOperatorRecord(unsigned char ptg);
			public: void PostBlobRead(BlobView* pBlobView);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgNumRecord : public ParsedExpressionRecord
		{
			public: PtgNumRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 9;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgNum;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: double m_value;
			public: PtgNumRecord(double value);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgMissArgRecord : public ParsedExpressionRecord
		{
			public: PtgMissArgRecord();
			public: PtgMissArgRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 1;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgMissArg;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgIntRecord : public ParsedExpressionRecord
		{
			public: PtgIntRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 3;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgInt;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned short m_integer;
			public: PtgIntRecord(unsigned short nInt);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgFuncVarRecord : public ParsedExpressionRecord
		{
			public: PtgFuncVarRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgFuncVar;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_cparams;
			protected: unsigned short m_tab;
			protected: unsigned short m_fCeFunc;
			public: PtgFuncVarRecord(unsigned short iftab, unsigned char cparams);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgFuncRecord : public ParsedExpressionRecord
		{
			public: PtgFuncRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 3;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgFunc;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved0;
			protected: unsigned short m_iftab;
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgBoolRecord : public ParsedExpressionRecord
		{
			public: PtgBoolRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgBool;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_boolean;
			public: PtgBoolRecord(bool bBool);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSumRecord : public ParsedExpressionRecord
		{
			public: PtgAttrSumRecord();
			public: PtgAttrSumRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgAttrSum;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_bitIf;
			protected: unsigned char m_reserved3;
			protected: unsigned short m_unused;
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSpaceRecord : public ParsedExpressionRecord
		{
			public: PtgAttrSpaceRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgAttrSpace;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_bitSpace;
			protected: unsigned char m_reserved3;
			protected: PtgAttrSpaceTypeStruct* m_type = 0;
			public: PtgAttrSpaceRecord(unsigned char nType, unsigned char nCount);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgAttrSpaceRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSemiRecord : public ParsedExpressionRecord
		{
			public: PtgAttrSemiRecord();
			public: PtgAttrSemiRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgAttrSemi;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_bitSemi;
			protected: unsigned char m_reserved2;
			protected: unsigned short m_unused;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrIfRecord : public ParsedExpressionRecord
		{
			public: PtgAttrIfRecord();
			public: PtgAttrIfRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgAttrIf;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_bitIf;
			protected: unsigned char m_reserved3;
			protected: unsigned short m_offset;
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrGotoRecord : public ParsedExpressionRecord
		{
			public: PtgAttrGotoRecord();
			public: PtgAttrGotoRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgAttrGoto;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_reserved0;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_bitGoto;
			protected: unsigned char m_reserved3;
			protected: unsigned short m_unused;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAreaRecord : public ParsedExpressionRecord
		{
			public: PtgAreaRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 9;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgArea;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved;
			protected: RgceAreaStruct* m_area = 0;
			public: PtgAreaRecord(Area* pArea);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgAreaRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgArea3dRecord : public ParsedExpressionRecord
		{
			public: PtgArea3dRecord(BlobView* pBlobView);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 11;
			protected: static const ParsedExpressionRecord::Type TYPE = ParsedExpressionRecord::Type::TYPE_PtgArea3d;
			protected: void SetDefaults();
			protected: unsigned char m_ptg;
			protected: unsigned char m_type;
			protected: unsigned char m_reserved;
			protected: unsigned short m_ixti;
			protected: RgceAreaStruct* m_area = 0;
			public: PtgArea3dRecord(Area3d* pArea3d, WorkbookGlobals* pWorkbookGlobals);
			public: virtual Token* GetToken(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~PtgArea3dRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Hsl
		{
			public: double m_fH;
			public: double m_fS;
			public: double m_fL;
			public: void ToColor(Color* pColor);
			public: void Set(const Color* pColor);
			protected: double max(double a, double b);
			protected: double min(double a, double b);
			protected: double hue2rgb(double p, double q, double t);
		};
		class XFExt : public BiffRecord
		{
			public: XFExt(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_XF_EXT;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_ixfe;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_cexts;
			protected: Color* m_pTempColor = 0;
			protected: OwnedVector<ExtPropStruct*>* m_pExtPropVector = 0;
			public: XFExt(unsigned short nXFIndex, const Color* pForegroundColor, const Color* pBackgroundColor, const Color* pTextColor);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetXFIndex();
			public: const Color* GetForegroundColor(Theme* pTheme);
			public: const Color* GetTextColor(Theme* pTheme);
			protected: ExtPropStruct* GetExtPropByType(unsigned short nType);
			protected: const Color* GetColorByType(unsigned short nType, Theme* pTheme);
			public: virtual ~XFExt();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XFCRC : public BiffRecord
		{
			public: XFCRC(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_XF_CRC;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned short m_reserved;
			protected: unsigned short m_cxfs;
			protected: unsigned int m_crc;
			public: XFCRC(OwnedVector<XF*>* pXFVector);
			public: bool ValidateCrc(Vector<XF*>* pXFVector);
			protected: unsigned int ComputateCrc(OwnedVector<XF*>* pXFVector);
			protected: unsigned int ComputateCrc(Vector<XF*>* pXFVector);
			public: virtual ~XFCRC();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XF : public BiffRecord
		{
			public: XF(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_XF;
			protected: void SetDefaults();
			protected: unsigned short m_ifnt;
			protected: unsigned short m_ifmt;
			protected: unsigned short m_fLocked;
			protected: unsigned short m_fHidden;
			protected: unsigned short m_fStyle;
			protected: unsigned short m_f123Prefix;
			protected: unsigned short m_ixfParent;
			protected: unsigned char m_alc;
			protected: unsigned char m_fWrap;
			protected: unsigned char m_alcV;
			protected: unsigned char m_fJustLast;
			protected: unsigned char m_trot;
			protected: unsigned char m_cIndent;
			protected: unsigned char m_fShrinkToFit;
			protected: unsigned char m_reserved1;
			protected: unsigned char m_iReadOrder;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_fAtrNum;
			protected: unsigned char m_fAtrFnt;
			protected: unsigned char m_fAtrAlc;
			protected: unsigned char m_fAtrBdr;
			protected: unsigned char m_fAtrPat;
			protected: unsigned char m_fAtrProt;
			protected: unsigned int m_dgLeft;
			protected: unsigned int m_dgRight;
			protected: unsigned int m_dgTop;
			protected: unsigned int m_dgBottom;
			protected: unsigned int m_icvLeft;
			protected: unsigned int m_icvRight;
			protected: unsigned int m_grbitDiag;
			protected: unsigned int m_icvTop;
			protected: unsigned int m_icvBottom;
			protected: unsigned int m_icvDiag;
			protected: unsigned int m_dgDiag;
			protected: unsigned int m_fHasXFExt;
			protected: unsigned int m_fls;
			protected: unsigned short m_icvFore;
			protected: unsigned short m_icvBack;
			protected: unsigned short m_fsxButton;
			protected: unsigned short m_reserved3;
			public: XF(unsigned short ifnt, unsigned short ifmt, unsigned char fStyle, unsigned short ixfParent, unsigned char fAtrNum, unsigned char fAtrFnt, unsigned char fAtrAlc, unsigned char fAtrBdr, unsigned char fAtrPat, unsigned char fAtrProt, unsigned char dgLeft, unsigned char dgRight, unsigned char dgTop, unsigned char dgBottom, unsigned char icvLeft, unsigned char icvRight, unsigned char icvTop, unsigned char icvBottom, unsigned char fHasXFExt, unsigned char fls, unsigned char icvFore, unsigned char icvBack);
			public: XF(unsigned short nFontIndex, unsigned short nFormatIndex, Style* pStyle, bool bHasXFExt);
			public: unsigned short GetFontIndex();
			public: unsigned short GetFormatIndex();
			public: unsigned short GetBackgroundPaletteIndex();
			public: Style::FillPattern GetFillPattern();
			public: unsigned short GetForegroundPaletteIndex();
			public: unsigned char GetHorizontalAlign();
			public: unsigned char GetVerticalAlign();
			protected: Line::Type BorderStyleToLineType(unsigned char nBorderStyle);
			protected: unsigned char LineTypeToBorderStyle(Line::Type eLineType);
			public: Line::Type GetTopBorderType();
			public: Line::Type GetRightBorderType();
			public: Line::Type GetBottomBorderType();
			public: Line::Type GetLeftBorderType();
			public: unsigned short GetTopBorderPaletteIndex();
			public: unsigned short GetRightBorderPaletteIndex();
			public: unsigned short GetBottomBorderPaletteIndex();
			public: unsigned short GetLeftBorderPaletteIndex();
			public: bool GetHasXFExt();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WsBoolRecord : public BiffRecord
		{
			public: WsBoolRecord();
			public: WsBoolRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_WSBOOL;
			protected: void SetDefaults();
			protected: unsigned short m_haxFlags;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WriteAccess : public BiffRecord
		{
			public: WriteAccess(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 3;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_WRITE_ACCESS;
			protected: void SetDefaults();
			protected: XLUnicodeStringStruct* m_userName = 0;
			protected: static const unsigned int FORCE_SIZE = 112;
			public: WriteAccess();
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~WriteAccess();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Window2Record : public BiffRecord
		{
			public: Window2Record(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_WINDOW2;
			protected: void SetDefaults();
			protected: unsigned short m_fDspFmlaRt;
			protected: unsigned short m_fDspGridRt;
			protected: unsigned short m_fDspRwColRt;
			protected: unsigned short m_fFrozenRt;
			protected: unsigned short m_fDspZerosRt;
			protected: unsigned short m_fDefaultHdr;
			protected: unsigned short m_fRightToLeft;
			protected: unsigned short m_fDspGuts;
			protected: unsigned short m_fFrozenNoSplit;
			protected: unsigned short m_fSelected;
			protected: unsigned short m_fPaged;
			protected: unsigned short m_fSLV;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_rwTop;
			protected: unsigned short m_colLeft;
			protected: unsigned short m_icvHdr;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_wScaleSLV;
			protected: unsigned short m_wScaleNormal;
			protected: unsigned short m_unused;
			protected: unsigned short m_reserved3;
			public: Window2Record(bool fSelected, bool fPaged, bool bShowGridlines);
			public: bool GetShowGridlines();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Window1 : public BiffRecord
		{
			public: Window1(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_WINDOW1;
			protected: void SetDefaults();
			protected: short m_xWn;
			protected: short m_yWn;
			protected: short m_dxWn;
			protected: short m_dyWn;
			protected: unsigned short m_fHidden;
			protected: unsigned short m_fIconic;
			protected: unsigned short m_fVeryHidden;
			protected: unsigned short m_fDspHScroll;
			protected: unsigned short m_fDspVScroll;
			protected: unsigned short m_fBotAdornment;
			protected: unsigned short m_fNoAFDateGroup;
			protected: unsigned short m_reserved;
			protected: unsigned short m_itabCur;
			protected: unsigned short m_itabFirst;
			protected: unsigned short m_ctabSel;
			protected: unsigned short m_wTabRatio;
			public: Window1(unsigned short nTabRatio);
			public: unsigned short GetTabRatio();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WinProtect : public BiffRecord
		{
			public: WinProtect();
			public: WinProtect(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_WIN_PROTECT;
			protected: void SetDefaults();
			protected: unsigned short m_fLockWn;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ValueRangeRecord : public BiffRecord
		{
			public: ValueRangeRecord();
			public: ValueRangeRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 42;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ValueRange;
			protected: void SetDefaults();
			protected: double m_numMin;
			protected: double m_numMax;
			protected: double m_numMajor;
			protected: double m_numMinor;
			protected: double m_numCross;
			protected: unsigned short m_fAutoMin;
			protected: unsigned short m_fAutoMax;
			protected: unsigned short m_fAutoMajor;
			protected: unsigned short m_fAutoMinor;
			protected: unsigned short m_fAutoCross;
			protected: unsigned short m_fLog;
			protected: unsigned short m_fReversed;
			protected: unsigned short m_fMaxCross;
			protected: unsigned short m_unused;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class VCenterRecord : public BiffRecord
		{
			public: VCenterRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_VCENTER;
			protected: void SetDefaults();
			protected: unsigned short m_vcenter;
			public: VCenterRecord(bool bCenter);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class UnitsRecord : public BiffRecord
		{
			public: UnitsRecord();
			public: UnitsRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Units;
			protected: void SetDefaults();
			protected: unsigned short m_unused;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TopMarginRecord : public BiffRecord
		{
			public: TopMarginRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_TOP_MARGIN;
			protected: void SetDefaults();
			protected: double m_num;
			public: TopMarginRecord(double num);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TickRecord : public BiffRecord
		{
			public: TickRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 30;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Tick;
			protected: void SetDefaults();
			protected: unsigned char m_tktMajor;
			protected: unsigned char m_tktMinor;
			protected: unsigned char m_tlt;
			protected: unsigned char m_wBkgMode;
			protected: unsigned int m_rgb;
			protected: unsigned int m_reserved1;
			protected: unsigned int m_reserved2;
			protected: unsigned int m_reserved3;
			protected: unsigned int m_reserved4;
			protected: unsigned short m_fAutoCo;
			protected: unsigned short m_fAutoMode;
			protected: unsigned short m_rot;
			protected: unsigned short m_fAutoRot;
			protected: unsigned short m_unused;
			protected: unsigned short m_iReadingOrder;
			protected: unsigned short m_icv;
			protected: unsigned short m_trot;
			public: TickRecord(bool fAutoRot);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Theme : public BiffRecord
		{
			public: Theme(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_THEME;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned int m_dwThemeVersion;
			public: Vector<unsigned int>* m_nColorVector = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned int GetColorByIndex(int nIndex);
			public: virtual ~Theme();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TextRecord : public BiffRecord
		{
			public: TextRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 32;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Text;
			protected: void SetDefaults();
			protected: unsigned char m_at;
			protected: unsigned char m_vat;
			protected: unsigned short m_wBkgMode;
			protected: unsigned int m_rgbText;
			protected: int m_x;
			protected: int m_y;
			protected: int m_dx;
			protected: int m_dy;
			protected: unsigned short m_fAutoColor;
			protected: unsigned short m_fShowKey;
			protected: unsigned short m_fShowValue;
			protected: unsigned short m_unused1;
			protected: unsigned short m_fAutoText;
			protected: unsigned short m_fGenerated;
			protected: unsigned short m_fDeleted;
			protected: unsigned short m_fAutoMode;
			protected: unsigned short m_unused2;
			protected: unsigned short m_fShowLabelAndPerc;
			protected: unsigned short m_fShowPercent;
			protected: unsigned short m_fShowBubbleSizes;
			protected: unsigned short m_fShowLabel;
			protected: unsigned short m_reserved;
			protected: unsigned short m_icvText;
			protected: unsigned short m_dlp;
			protected: unsigned short m_unused3;
			protected: unsigned short m_iReadingOrder;
			protected: unsigned short m_trot;
			public: TextRecord(int x, int y, unsigned int rgbText, bool fAutoColor, bool fAutoText, bool fGenerated, bool fAutoMode, unsigned short icvText, unsigned short trot);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SupBookRecord : public BiffRecord
		{
			public: SupBookRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SUP_BOOK;
			protected: void SetDefaults();
			protected: unsigned short m_ctab;
			protected: unsigned short m_cch;
			public: SupBookRecord(unsigned short nNumWorksheet);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StyleRecord : public BiffRecord
		{
			public: StyleRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_STYLE;
			protected: void SetDefaults();
			protected: unsigned short m_ixfe;
			protected: unsigned short m_unused;
			protected: unsigned short m_fBuiltIn;
			protected: BuiltInStyleStruct* m_builtInData = 0;
			protected: XLUnicodeStringStruct* m_user = 0;
			public: StyleRecord(unsigned short nXFIndex, const char* szName);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~StyleRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StartBlockRecord : public BiffRecord
		{
			public: StartBlockRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 12;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_StartBlock;
			protected: void SetDefaults();
			protected: FrtHeaderOldStruct* m_frtHeaderOld = 0;
			protected: unsigned short m_iObjectKind;
			protected: unsigned short m_iObjectContext;
			protected: unsigned short m_iObjectInstance1;
			protected: unsigned short m_iObjectInstance2;
			public: StartBlockRecord(unsigned short iObjectKind, unsigned short iObjectContext, unsigned short iObjectInstance1, unsigned short iObjectInstance2);
			protected: void PostSetDefaults();
			public: virtual ~StartBlockRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SstRecord : public BiffRecord
		{
			public: SstRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SST;
			protected: void SetDefaults();
			protected: int m_cstTotal;
			protected: int m_cstUnique;
			protected: Blob* m_pHaxBlob = 0;
			protected: OwnedVector<XLUnicodeRichExtendedString*>* m_pRgb = 0;
			public: SstRecord(SharedStringContainer* pSharedStringContainer);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: int GetNumString();
			public: const char* GetString(int nIndex);
			public: virtual ~SstRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShtPropsRecord : public BiffRecord
		{
			public: ShtPropsRecord();
			public: ShtPropsRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ShtProps;
			protected: void SetDefaults();
			protected: unsigned short m_fManSerAlloc;
			protected: unsigned short m_fPlotVisOnly;
			protected: unsigned short m_fNotSizeWith;
			protected: unsigned short m_fManPlotArea;
			protected: unsigned short m_fAlwaysAutoPlotArea;
			protected: unsigned short m_reserved1;
			protected: unsigned char m_mdBlank;
			protected: unsigned char m_reserved2;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShapePropsStreamRecord : public BiffRecord
		{
			public: ShapePropsStreamRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 24;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ShapePropsStream;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned short m_wObjContext;
			protected: unsigned short m_unused;
			protected: unsigned int m_dwChecksum;
			protected: unsigned int m_cb;
			protected: InternalString* m_rgb = 0;
			public: ShapePropsStreamRecord(unsigned short wObjContext, unsigned short unused, unsigned int dwChecksum);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			public: virtual ~ShapePropsStreamRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SetupRecord : public BiffRecord
		{
			public: SetupRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 34;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SETUP;
			protected: void SetDefaults();
			protected: unsigned short m_iPaperSize;
			protected: unsigned short m_iScale;
			protected: short m_iPageStart;
			protected: unsigned short m_iFitWidth;
			protected: unsigned short m_iFitHeight;
			protected: unsigned short m_fLeftToRight;
			protected: unsigned short m_fPortrait;
			protected: unsigned short m_fNoPls;
			protected: unsigned short m_fNoColor;
			protected: unsigned short m_fDraft;
			protected: unsigned short m_fNotes;
			protected: unsigned short m_fNoOrient;
			protected: unsigned short m_fUsePage;
			protected: unsigned short m_unused1;
			protected: unsigned short m_fEndNotes;
			protected: unsigned short m_iErrors;
			protected: unsigned short m_reserved;
			protected: unsigned short m_iRes;
			protected: unsigned short m_iVRes;
			protected: double m_numHdr;
			protected: double m_numFtr;
			protected: unsigned short m_iCopies;
			public: SetupRecord(bool bPortrait);
			public: bool GetPortrait();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesTextRecord : public BiffRecord
		{
			public: SeriesTextRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SeriesText;
			protected: void SetDefaults();
			protected: unsigned short m_reserved;
			protected: ShortXLUnicodeStringStruct* m_stText = 0;
			public: SeriesTextRecord(const char* szText);
			public: virtual ~SeriesTextRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesRecord : public BiffRecord
		{
			public: SeriesRecord();
			public: SeriesRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 12;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Series;
			protected: void SetDefaults();
			protected: unsigned short m_sdtX;
			protected: unsigned short m_sdtY;
			protected: unsigned short m_cValx;
			protected: unsigned short m_cValy;
			protected: unsigned short m_sdtBSize;
			protected: unsigned short m_cValBSize;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SerToCrtRecord : public BiffRecord
		{
			public: SerToCrtRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SerToCrt;
			protected: void SetDefaults();
			protected: unsigned short m_id;
			public: SerToCrtRecord(unsigned short id);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SelectionRecord : public BiffRecord
		{
			public: SelectionRecord();
			public: SelectionRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 15;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SELECTION;
			protected: void SetDefaults();
			protected: unsigned char m_pnn;
			protected: unsigned short m_rwAct;
			protected: unsigned short m_colAct;
			protected: unsigned short m_irefAct;
			protected: unsigned short m_cref;
			protected: unsigned short m_rwFirst;
			protected: unsigned short m_rwLast;
			protected: unsigned char m_colFirst;
			protected: unsigned char m_colLast;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SclRecord : public BiffRecord
		{
			public: SclRecord();
			public: SclRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SCL;
			protected: void SetDefaults();
			protected: short m_nscl;
			protected: short m_dscl;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ScatterRecord : public BiffRecord
		{
			public: ScatterRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 6;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Scatter;
			protected: void SetDefaults();
			protected: unsigned short m_pcBubbleSizeRatio;
			protected: unsigned short m_wBubbleSize;
			protected: unsigned short m_fBubbles;
			protected: unsigned short m_fShowNegBubbles;
			protected: unsigned short m_fHasShadow;
			protected: unsigned short m_reserved;
			public: ScatterRecord(Chart::Type eType);
			public: Chart::Type GetChartType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SIIndexRecord : public BiffRecord
		{
			public: SIIndexRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_SIIndex;
			protected: void SetDefaults();
			protected: unsigned short m_numIndex;
			public: SIIndexRecord(unsigned short numIndex);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RowRecord : public BiffRecord
		{
			public: RowRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ROW;
			protected: void SetDefaults();
			protected: RwStruct* m_rw = 0;
			protected: unsigned short m_colMic;
			protected: unsigned short m_colMac;
			protected: unsigned short m_miyRw;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_unused1;
			protected: unsigned char m_iOutLevel;
			protected: unsigned char m_reserved2;
			protected: unsigned char m_fCollapsed;
			protected: unsigned char m_fDyZero;
			protected: unsigned char m_fUnsynced;
			protected: unsigned char m_fGhostDirty;
			protected: unsigned char m_reserved3;
			protected: unsigned short m_ixfe_val;
			protected: unsigned short m_fExAsc;
			protected: unsigned short m_fExDes;
			protected: unsigned short m_fPhonetic;
			protected: unsigned short m_unused2;
			public: bool m_bTopMedium;
			public: bool m_bTopThick;
			public: bool m_bBottomMedium;
			public: bool m_bBottomThick;
			public: RowRecord(unsigned short nRow, unsigned short nHeight);
			protected: void PostSetDefaults();
			public: unsigned short GetRow();
			public: unsigned short GetHeight();
			public: void SetHeight(unsigned short nHeight);
			public: void SetTopMedium();
			public: void SetTopThick();
			public: void SetBottomMedium();
			public: void SetBottomThick();
			public: bool GetBottomThick();
			public: virtual ~RowRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RkRecord : public BiffRecord
		{
			public: RkRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 10;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_RK;
			protected: void SetDefaults();
			protected: unsigned short m_nRowIndex;
			protected: unsigned short m_nColumnIndex;
			protected: unsigned short m_nXfIndex;
			protected: unsigned int m_nRkValue;
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: unsigned int GetRkValue();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RightMarginRecord : public BiffRecord
		{
			public: RightMarginRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_RIGHT_MARGIN;
			protected: void SetDefaults();
			protected: double m_num;
			public: RightMarginRecord(double num);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RefreshAllRecord : public BiffRecord
		{
			public: RefreshAllRecord();
			public: RefreshAllRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_REFRESH_ALL;
			protected: void SetDefaults();
			protected: unsigned short m_refreshAll;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RRTabId : public BiffRecord
		{
			public: RRTabId(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_RR_TAB_ID;
			protected: void SetDefaults();
			public: unsigned short m_nNumWorksheet;
			public: RRTabId(unsigned short nNumWorksheet);
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ProtectRecord : public BiffRecord
		{
			public: ProtectRecord();
			public: ProtectRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PROTECT;
			protected: void SetDefaults();
			protected: unsigned short m_fLock;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Prot4RevRecord : public BiffRecord
		{
			public: Prot4RevRecord();
			public: Prot4RevRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PROT_4_REV;
			protected: void SetDefaults();
			protected: unsigned short m_fRevLock;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Prot4RevPassRecord : public BiffRecord
		{
			public: Prot4RevPassRecord();
			public: Prot4RevPassRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PROT_4_REV_PASS;
			protected: void SetDefaults();
			protected: unsigned short m_protPwdRev;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintSizeRecord : public BiffRecord
		{
			public: PrintSizeRecord();
			public: PrintSizeRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PRINT_SIZE;
			protected: void SetDefaults();
			protected: unsigned short m_printSize;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintRowColRecord : public BiffRecord
		{
			public: PrintRowColRecord();
			public: PrintRowColRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PrintRowCol;
			protected: void SetDefaults();
			protected: unsigned short m_printRwCol;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintGridRecord : public BiffRecord
		{
			public: PrintGridRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PrintGrid;
			protected: void SetDefaults();
			protected: unsigned short m_fPrintGrid;
			protected: unsigned short m_unused;
			public: PrintGridRecord(bool bPrintGridlines);
			public: bool GetPrintGridlines();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PosRecord : public BiffRecord
		{
			public: PosRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Pos;
			protected: void SetDefaults();
			protected: unsigned short m_mdTopLt;
			protected: unsigned short m_mdBotRt;
			protected: short m_x1;
			protected: unsigned short m_unused1;
			protected: short m_y1;
			protected: unsigned short m_unused2;
			protected: short m_x2;
			protected: unsigned short m_unused3;
			protected: short m_y2;
			protected: unsigned short m_unused4;
			public: PosRecord(short x1, short y1, unsigned short unused2, short x2, short y2);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PlotGrowthRecord : public BiffRecord
		{
			public: PlotGrowthRecord();
			public: PlotGrowthRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PlotGrowth;
			protected: void SetDefaults();
			protected: int m_dxPlotGrowth;
			protected: int m_dyPlotGrowth;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PlotAreaRecord : public BiffRecord
		{
			public: PlotAreaRecord();
			public: PlotAreaRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PlotArea;
			protected: void SetDefaults();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PieFormatRecord : public BiffRecord
		{
			public: PieFormatRecord();
			public: PieFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PieFormat;
			protected: void SetDefaults();
			protected: unsigned short m_pcExplode;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PasswordRecord : public BiffRecord
		{
			public: PasswordRecord();
			public: PasswordRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PASSWORD;
			protected: void SetDefaults();
			protected: unsigned short m_wPassword;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PaletteRecord : public BiffRecord
		{
			public: PaletteRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_PALETTE;
			protected: void SetDefaults();
			protected: unsigned short m_ccv;
			protected: unsigned int m_rgColor[BiffWorkbookGlobals::NUM_CUSTOM_PALETTE_ENTRY];
			public: PaletteRecord();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostSetDefaults();
			public: unsigned int GetColorByIndex(unsigned short nIndex);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ObjectLinkRecord : public BiffRecord
		{
			public: ObjectLinkRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 6;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ObjectLink;
			protected: void SetDefaults();
			protected: unsigned short m_wLinkObj;
			protected: unsigned short m_wLinkVar1;
			protected: unsigned short m_wLinkVar2;
			public: ObjectLinkRecord(unsigned short nLinkObject, unsigned short nSeriesIndex, unsigned short nCategoryIndex);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ObjRecord : public BiffRecord
		{
			public: ObjRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 22;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_OBJ;
			protected: void SetDefaults();
			protected: FtCmoStruct* m_cmo = 0;
			protected: FtCfStruct* m_pictFormat = 0;
			protected: FtPioGrbitStruct* m_pictFlags = 0;
			public: ObjRecord(unsigned short nIndex, FtCmoStruct::ObjType eType);
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			protected: void PostSetDefaults();
			public: FtCmoStruct::ObjType GetType();
			public: virtual ~ObjRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class NumberRecord : public BiffRecord
		{
			public: NumberRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 14;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_NUMBER;
			protected: void SetDefaults();
			protected: CellStruct* m_cell = 0;
			protected: double m_num;
			public: NumberRecord(unsigned short nX, unsigned short nY, unsigned short nXfIndex, double fNumber);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: double GetNumber();
			public: virtual ~NumberRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MulRkRecord : public BiffRecord
		{
			public: MulRkRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MULRK;
			protected: void SetDefaults();
			protected: RwStruct* m_rw = 0;
			protected: ColStruct* m_col = 0;
			protected: OwnedVector<RkRecStruct*>* m_pRkRecVector = 0;
			protected: unsigned short m_colLast;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetNumRk();
			public: unsigned short GetXfIndexByIndex(unsigned short nIndex);
			public: unsigned int GetRkValueByIndex(unsigned short nIndex);
			public: virtual ~MulRkRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MulBlank : public BiffRecord
		{
			public: MulBlank(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MULBLANK;
			protected: void SetDefaults();
			protected: RwStruct* m_rw = 0;
			protected: ColStruct* m_col = 0;
			protected: OwnedVector<IXFCellStruct*>* m_pIXFCellVector = 0;
			protected: unsigned short m_colLast;
			public: MulBlank(unsigned short nX, unsigned short nY, Vector<int>* pXfIndexVector);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetNumColumn();
			public: unsigned short GetXfIndexByIndex(unsigned short nIndex);
			public: virtual ~MulBlank();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MsoDrawingRecord_Position
		{
			public: unsigned short m_nCellX1;
			public: unsigned short m_nSubCellX1;
			public: unsigned short m_nCellY1;
			public: unsigned short m_nSubCellY1;
			public: unsigned short m_nCellX2;
			public: unsigned short m_nSubCellX2;
			public: unsigned short m_nCellY2;
			public: unsigned short m_nSubCellY2;
		};
		class MsoDrawingRecord : public BiffRecord
		{
			public: MsoDrawingRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MSO_DRAWING;
			protected: void SetDefaults();
			protected: OfficeArtRecord* m_pOfficeArtRecord = 0;
			protected: MsoDrawingRecord_Position* m_pPosition = 0;
			protected: Blob* m_pBlob = 0;
			public: MsoDrawingRecord(OfficeArtDgContainerRecord* pOfficeArtDgContainerRecord, unsigned int nWriteSize);
			public: MsoDrawingRecord(OfficeArtSpContainerRecord* pOfficeArtSpContainerRecord);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: const MsoDrawingRecord_Position* GetPosition();
			public: OfficeArtFOPTEStruct* GetProperty(OfficeArtRecord::OPIDType eType);
			public: virtual ~MsoDrawingRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MsoDrawingGroupRecord : public BiffRecord
		{
			public: MsoDrawingGroupRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MSO_DRAWING_GROUP;
			protected: void SetDefaults();
			protected: OfficeArtDggContainerRecord* m_pOfficeArtDggContainerRecord = 0;
			public: MsoDrawingGroupRecord(Vector<Picture*>* pPictureVector);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: OfficeArtDggContainerRecord* GetOfficeArtDggContainerRecord();
			public: virtual ~MsoDrawingGroupRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Mms : public BiffRecord
		{
			public: Mms();
			public: Mms(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MMS;
			protected: void SetDefaults();
			protected: unsigned char m_reserved1;
			protected: unsigned char m_reserved2;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MergeCellsRecord : public BiffRecord
		{
			public: MergeCellsRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MergeCells;
			protected: void SetDefaults();
			protected: unsigned short m_cmcs;
			protected: OwnedVector<Ref8Struct*>* m_pRef8Vector = 0;
			public: MergeCellsRecord(OwnedVector<MergedCell*>* pMergedCellVector, int nOffset, int nSize);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetNumMergedCell();
			public: const Ref8Struct* GetMergedCell(unsigned short nIndex);
			public: virtual ~MergeCellsRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MarkerFormatRecord : public BiffRecord
		{
			public: MarkerFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_MarkerFormat;
			protected: void SetDefaults();
			protected: unsigned int m_rgbFore;
			protected: unsigned int m_rgbBack;
			protected: unsigned short m_imk;
			protected: unsigned short m_fAuto;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_fNotShowInt;
			protected: unsigned short m_fNotShowBrd;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_icvFore;
			protected: unsigned short m_icvBack;
			protected: unsigned int m_miSize;
			public: MarkerFormatRecord(Marker* pMarker);
			public: void ModifyMarker(Marker* pMarker, BiffWorkbookGlobals* pBiffWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineRecord : public BiffRecord
		{
			public: LineRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Line;
			protected: void SetDefaults();
			protected: unsigned short m_fStacked;
			protected: unsigned short m_f100;
			protected: unsigned short m_fHasShadow;
			protected: unsigned short m_reserved;
			public: LineRecord(Chart::Type eType);
			public: Chart::Type GetChartType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineFormatRecord : public BiffRecord
		{
			public: LineFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 12;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_LineFormat;
			protected: void SetDefaults();
			protected: unsigned int m_rgb;
			protected: unsigned short m_lns;
			protected: short m_we;
			protected: unsigned short m_fAuto;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_fAxisOn;
			protected: unsigned short m_fAutoCo;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_icv;
			public: LineFormatRecord(unsigned int rgb, unsigned short lns, short we, bool fAuto, bool fAxisOn, bool fAutoCo, unsigned short icv);
			public: LineFormatRecord(Line* pLine, bool bAxis);
			public: void ModifyLine(Line* pLine, BiffWorkbookGlobals* pBiffWorkbookGlobals);
			public: void PackForChecksum(BlobView* pBlobView);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LegendRecord : public BiffRecord
		{
			public: LegendRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Legend;
			protected: void SetDefaults();
			protected: unsigned int m_x;
			protected: unsigned int m_y;
			protected: unsigned int m_dx;
			protected: unsigned int m_dy;
			protected: unsigned char m_unused;
			protected: unsigned char m_wSpace;
			protected: unsigned short m_fAutoPosition;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_fAutoPosX;
			protected: unsigned short m_fAutoPosY;
			protected: unsigned short m_fVert;
			protected: unsigned short m_fWasDataTable;
			protected: unsigned short m_reserved2;
			public: LegendRecord(const Legend* pLegend);
			public: void ModifyLegend(Legend* pLegend, BiffWorkbookGlobals* pBiffWorkbookGlobals);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LeftMarginRecord : public BiffRecord
		{
			public: LeftMarginRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_LEFT_MARGIN;
			protected: void SetDefaults();
			protected: double m_num;
			public: LeftMarginRecord(double num);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LabelSstRecord : public BiffRecord
		{
			public: LabelSstRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 10;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_LABELSST;
			protected: void SetDefaults();
			protected: unsigned short m_nRowIndex;
			protected: unsigned short m_nColumnIndex;
			protected: unsigned short m_nXfIndex;
			protected: unsigned int m_nSstIndex;
			public: LabelSstRecord(unsigned short nX, unsigned short nY, unsigned short nXfIndex, unsigned int nSstIndex);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: unsigned int GetSstIndex();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class InterfaceHdr : public BiffRecord
		{
			public: InterfaceHdr();
			public: InterfaceHdr(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_INTERFACE_HDR;
			protected: void SetDefaults();
			protected: unsigned short m_codePage;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class InterfaceEnd : public BiffRecord
		{
			public: InterfaceEnd();
			public: InterfaceEnd(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_INTERFACE_END;
			protected: void SetDefaults();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HideObj : public BiffRecord
		{
			public: HideObj();
			public: HideObj(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_HIDE_OBJ;
			protected: void SetDefaults();
			protected: unsigned short m_hideObj;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HeaderFooterRecord : public BiffRecord
		{
			public: HeaderFooterRecord();
			public: HeaderFooterRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 38;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_HeaderFooter;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned char m_guidSView_0;
			protected: unsigned char m_guidSView_1;
			protected: unsigned char m_guidSView_2;
			protected: unsigned char m_guidSView_3;
			protected: unsigned char m_guidSView_4;
			protected: unsigned char m_guidSView_5;
			protected: unsigned char m_guidSView_6;
			protected: unsigned char m_guidSView_7;
			protected: unsigned char m_guidSView_8;
			protected: unsigned char m_guidSView_9;
			protected: unsigned char m_guidSView_10;
			protected: unsigned char m_guidSView_11;
			protected: unsigned char m_guidSView_12;
			protected: unsigned char m_guidSView_13;
			protected: unsigned char m_guidSView_14;
			protected: unsigned char m_guidSView_15;
			protected: unsigned short m_fHFDiffOddEven;
			protected: unsigned short m_fHFDiffFirst;
			protected: unsigned short m_fHFScaleWithDoc;
			protected: unsigned short m_fHFAlignMargins;
			protected: unsigned short m_unused;
			protected: unsigned short m_cchHeaderEven;
			protected: unsigned short m_cchFooterEven;
			protected: unsigned short m_cchHeaderFirst;
			protected: unsigned short m_cchFooterFirst;
			protected: void PostSetDefaults();
			public: virtual ~HeaderFooterRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HCenterRecord : public BiffRecord
		{
			public: HCenterRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_HCENTER;
			protected: void SetDefaults();
			protected: unsigned short m_hcenter;
			public: HCenterRecord(bool bCenter);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class GelFrameRecord : public BiffRecord
		{
			public: GelFrameRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_GelFrame;
			protected: void SetDefaults();
			protected: OfficeArtFOPTRecord* m_OPT1 = 0;
			protected: OfficeArtTertiaryFOPTRecord* m_OPT2 = 0;
			public: GelFrameRecord(Fill* pFill, WorkbookGlobals* pWorkbookGlobals);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: void ModifyFill(Fill* pFill, WorkbookGlobals* pWorkbookGlobals);
			public: void PackForChecksum(BlobView* pBlobView);
			public: virtual ~GelFrameRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrameRecord : public BiffRecord
		{
			public: FrameRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 4;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Frame;
			protected: void SetDefaults();
			protected: unsigned short m_frt;
			protected: unsigned short m_fAutoSize;
			protected: unsigned short m_fAutoPosition;
			protected: unsigned short m_reserved;
			public: FrameRecord(bool fAutoSize);
			public: bool GetDropShadow();
			public: bool GetAutoSize();
			public: bool GetAutoPosition();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FormulaRecord : public BiffRecord
		{
			public: FormulaRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_FORMULA;
			protected: void SetDefaults();
			protected: CellStruct* m_cell = 0;
			protected: FormulaValueStruct* m_val = 0;
			protected: unsigned short m_fAlwaysCalc;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_fFill;
			protected: unsigned short m_fShrFmla;
			protected: unsigned short m_reserved2;
			protected: unsigned short m_fClearErrors;
			protected: unsigned short m_reserved3;
			protected: unsigned int m_chn;
			protected: CellParsedFormulaStruct* m_formula = 0;
			public: FormulaRecord(unsigned short nX, unsigned short nY, unsigned short nXfIndex, Formula* pFormula, WorkbookGlobals* pWorkbookGlobals);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: Formula* GetFormula(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~FormulaRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Format : public BiffRecord
		{
			public: Format(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 5;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_FORMAT;
			protected: void SetDefaults();
			protected: unsigned short m_ifmt;
			protected: XLUnicodeStringStruct* m_stFormat = 0;
			public: Format(unsigned short ifmt, const char* szFormat);
			public: unsigned short GetFormatIndex();
			public: const char* GetFormat();
			public: virtual ~Format();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontXRecord : public BiffRecord
		{
			public: FontXRecord();
			public: FontXRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_FontX;
			protected: void SetDefaults();
			protected: unsigned short m_iFont;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontRecord : public BiffRecord
		{
			public: FontRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_FONT;
			protected: void SetDefaults();
			protected: unsigned short m_dyHeight;
			protected: unsigned short m_unused1;
			protected: unsigned short m_fItalic;
			protected: unsigned short m_unused2;
			protected: unsigned short m_fStrikeOut;
			protected: unsigned short m_fOutline;
			protected: unsigned short m_fShadow;
			protected: unsigned short m_fCondense;
			protected: unsigned short m_fExtend;
			protected: unsigned short m_reserved;
			protected: IcvFontStruct* m_icv = 0;
			protected: unsigned short m_bls;
			protected: unsigned short m_sss;
			protected: unsigned char m_uls;
			protected: unsigned char m_bFamily;
			protected: unsigned char m_bCharSet;
			protected: unsigned char m_unused3;
			protected: ShortXLUnicodeStringStruct* m_fontName = 0;
			public: FontRecord(const char* szFontName, unsigned short nSizeTwips, unsigned short nColourIndex, bool bBold, bool bItalic, Font::Underline eUnderline, unsigned char unused1 = 0, unsigned char unused3 = 181);
			public: unsigned short GetColourIndex();
			public: const char* GetName();
			public: unsigned short GetSizeTwips();
			public: bool GetBold();
			public: bool GetItalic();
			public: Font::Underline GetUnderline();
			public: virtual ~FontRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ExternSheetRecord : public BiffRecord
		{
			public: ExternSheetRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_EXTERN_SHEET;
			protected: void SetDefaults();
			protected: unsigned short m_cXTI;
			protected: OwnedVector<XTIStruct*>* m_pXTIVector = 0;
			public: ExternSheetRecord(OwnedVector<WorksheetRange*>* pWorksheetRangeVector);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetNumXTI();
			public: XTIStruct* GetXTIByIndex(unsigned short nIndex);
			public: virtual ~ExternSheetRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Excel9File : public BiffRecord
		{
			public: Excel9File();
			public: Excel9File(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_EXCEL9_FILE;
			protected: void SetDefaults();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class EndBlockRecord : public BiffRecord
		{
			public: EndBlockRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 12;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_EndBlock;
			protected: void SetDefaults();
			protected: FrtHeaderOldStruct* m_frtHeaderOld = 0;
			protected: unsigned short m_iObjectKind;
			protected: unsigned short m_unused1;
			protected: unsigned short m_unused2;
			protected: unsigned short m_unused3;
			public: EndBlockRecord(unsigned short iObjectKind);
			public: virtual ~EndBlockRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DimensionRecord : public BiffRecord
		{
			public: DimensionRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 14;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DIMENSION;
			protected: void SetDefaults();
			protected: unsigned int m_nFirstUsedRow;
			protected: unsigned int m_nLastUsedRow;
			protected: unsigned short m_nFirstUsedColumn;
			protected: unsigned short m_nLastUsedColumn;
			protected: unsigned short m_nUnused;
			public: DimensionRecord(unsigned int nFirstUsedRow, unsigned int nLastUsedRow, unsigned short nFirstUsedColumn, unsigned short nLastUsedColumn, unsigned short nUnused);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefaultTextRecord : public BiffRecord
		{
			public: DefaultTextRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DefaultText;
			protected: void SetDefaults();
			protected: unsigned short m_id;
			public: DefaultTextRecord(unsigned short id);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefaultRowHeight : public BiffRecord
		{
			public: DefaultRowHeight(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DEFAULTROWHEIGHT;
			protected: void SetDefaults();
			protected: unsigned short m_fUnsynced;
			protected: unsigned short m_fDyZero;
			protected: unsigned short m_fExAsc;
			protected: unsigned short m_fExDsc;
			protected: unsigned short m_reserved;
			protected: short m_miyRw;
			protected: short m_miyRwHidden;
			public: DefaultRowHeight(short nRowHeight);
			public: void PostBlobRead(BlobView* pBlobView);
			public: void PostBlobWrite(BlobView* pBlobView);
			public: short GetRowHeight();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefColWidthRecord : public BiffRecord
		{
			public: DefColWidthRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DEFCOLWIDTH;
			protected: void SetDefaults();
			protected: unsigned short m_cchdefColWidth;
			public: DefColWidthRecord(unsigned short cchdefColWidth);
			public: unsigned short GetColumnWidth();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Date1904 : public BiffRecord
		{
			public: Date1904();
			public: Date1904(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DATE_1904;
			protected: void SetDefaults();
			protected: unsigned short m_f1904DateSystem;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DataFormatRecord : public BiffRecord
		{
			public: DataFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DataFormat;
			protected: void SetDefaults();
			protected: unsigned short m_xi;
			protected: unsigned short m_yi;
			protected: unsigned short m_iss;
			protected: unsigned short m_fXL4iss;
			protected: unsigned short m_reserved;
			public: DataFormatRecord(unsigned short nIndex);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DSF : public BiffRecord
		{
			public: DSF();
			public: DSF(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_DSF;
			protected: void SetDefaults();
			protected: short m_reserved;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtMlFrtRecord : public BiffRecord
		{
			public: CrtMlFrtRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CrtMlFrt;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned int m_cb;
			protected: void PostSetDefaults();
			public: virtual ~CrtMlFrtRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtLinkRecord : public BiffRecord
		{
			public: CrtLinkRecord();
			public: CrtLinkRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 10;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CrtLink;
			protected: void SetDefaults();
			protected: unsigned short m_unused0;
			protected: unsigned short m_unused1;
			protected: unsigned short m_unused2;
			protected: unsigned short m_unused3;
			protected: unsigned short m_unused4;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtLayout12ARecord : public BiffRecord
		{
			public: CrtLayout12ARecord();
			public: CrtLayout12ARecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 68;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CrtLayout12A;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned int m_dwCheckSum;
			protected: unsigned short m_fLayoutTargetInner;
			protected: unsigned short m_reserved1;
			protected: short m_xTL;
			protected: short m_yTL;
			protected: short m_xBR;
			protected: short m_yBR;
			protected: unsigned short m_wXMode;
			protected: unsigned short m_wYMode;
			protected: unsigned short m_wWidthMode;
			protected: unsigned short m_wHeightMode;
			protected: double m_x;
			protected: double m_y;
			protected: double m_dx;
			protected: double m_dy;
			protected: unsigned short m_reserved2;
			public: void PostSetDefaults();
			public: virtual ~CrtLayout12ARecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColInfoRecord : public BiffRecord
		{
			public: ColInfoRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 10;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_COLINFO;
			protected: void SetDefaults();
			protected: unsigned short m_colFirst;
			protected: unsigned short m_colLast;
			protected: unsigned short m_coldx;
			protected: unsigned short m_ixfe;
			protected: unsigned short m_fHidden;
			protected: unsigned short m_fUserSet;
			protected: unsigned short m_fBestFit;
			protected: unsigned short m_fPhonetic;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_iOutLevel;
			protected: unsigned short m_unused1;
			protected: unsigned short m_fCollapsed;
			protected: unsigned short m_reserved2;
			public: ColInfoRecord(unsigned short nFirstColumn, unsigned short nLastColumn, unsigned short nColumnWidth, bool bHidden);
			public: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned short GetFirstColumn();
			public: unsigned short GetLastColumn();
			public: unsigned short GetColumnWidth();
			public: bool GetHidden();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CodePage : public BiffRecord
		{
			public: CodePage();
			public: CodePage(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CODE_PAGE;
			protected: void SetDefaults();
			protected: short m_cv;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartRecord : public BiffRecord
		{
			public: ChartRecord();
			public: ChartRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Chart;
			protected: void SetDefaults();
			protected: unsigned int m_x;
			protected: unsigned int m_y;
			protected: unsigned int m_dx;
			protected: unsigned int m_dy;
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetWidth();
			public: unsigned short GetHeight();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartFrtInfoRecord : public BiffRecord
		{
			public: ChartFrtInfoRecord();
			public: ChartFrtInfoRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 24;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ChartFrtInfo;
			protected: void SetDefaults();
			protected: FrtHeaderOldStruct* m_frtHeaderOld = 0;
			protected: unsigned char m_verOriginator;
			protected: unsigned char m_verWriter;
			protected: unsigned short m_cCFRTID;
			protected: unsigned short m_hax1;
			protected: unsigned short m_hax2;
			protected: unsigned short m_hax3;
			protected: unsigned short m_hax4;
			protected: unsigned short m_hax5;
			protected: unsigned short m_hax6;
			protected: unsigned short m_hax7;
			protected: unsigned short m_hax8;
			protected: void PostSetDefaults();
			public: virtual ~ChartFrtInfoRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartFormatRecord : public BiffRecord
		{
			public: ChartFormatRecord();
			public: ChartFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_ChartFormat;
			protected: void SetDefaults();
			protected: unsigned int m_reserved1;
			protected: unsigned int m_reserved2;
			protected: unsigned int m_reserved3;
			protected: unsigned int m_reserved4;
			protected: unsigned short m_fVaried;
			protected: unsigned short m_reserved5;
			protected: unsigned short m_icrt;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Chart3DBarShapeRecord : public BiffRecord
		{
			public: Chart3DBarShapeRecord();
			public: Chart3DBarShapeRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Chart3DBarShape;
			protected: void SetDefaults();
			protected: unsigned char m_riser;
			protected: unsigned char m_taper;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CatSerRangeRecord : public BiffRecord
		{
			public: CatSerRangeRecord();
			public: CatSerRangeRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CatSerRange;
			protected: void SetDefaults();
			protected: short m_catCross;
			protected: short m_catLabel;
			protected: short m_catMark;
			protected: unsigned short m_fBetween;
			protected: unsigned short m_fMaxCross;
			protected: unsigned short m_fReverse;
			protected: unsigned short m_reserved;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CatLabRecord : public BiffRecord
		{
			public: CatLabRecord();
			public: CatLabRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 12;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CatLab;
			protected: void SetDefaults();
			protected: FrtHeaderOldStruct* m_frtHeaderOld = 0;
			protected: unsigned short m_wOffset;
			protected: unsigned short m_at;
			protected: unsigned short m_cAutoCatLabelReal;
			protected: unsigned short m_unused;
			protected: unsigned short m_reserved;
			protected: void PostSetDefaults();
			public: virtual ~CatLabRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CalcPrecision : public BiffRecord
		{
			public: CalcPrecision();
			public: CalcPrecision(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CALC_PRECISION;
			protected: void SetDefaults();
			protected: unsigned short m_fFullPrec;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CalcCountRecord : public BiffRecord
		{
			public: CalcCountRecord();
			public: CalcCountRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_CALCCOUNT;
			protected: void SetDefaults();
			protected: unsigned short m_nMaxNumIteration;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BuiltInFnGroupCount : public BiffRecord
		{
			public: BuiltInFnGroupCount();
			public: BuiltInFnGroupCount(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BUILT_IN_FN_GROUP_COUNT;
			protected: void SetDefaults();
			protected: unsigned short m_count;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BraiRecord : public BiffRecord
		{
			public: BraiRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 6;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BRAI;
			protected: void SetDefaults();
			protected: unsigned char m_id;
			protected: unsigned char m_rt;
			protected: unsigned short m_fUnlinkedIfmt;
			protected: unsigned short m_reserved;
			protected: unsigned short m_ifmt;
			protected: CellParsedFormulaStruct* m_formula = 0;
			public: BraiRecord(unsigned char id, unsigned char rt, bool fUnlinkedIfmt, unsigned short ifmt, Formula* pFormula, WorkbookGlobals* pWorkbookGlobals);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: unsigned char GetId();
			public: unsigned char GetRt();
			public: Formula* GetFormula(WorkbookGlobals* pWorkbookGlobals);
			public: virtual ~BraiRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoundSheet8Record : public BiffRecord
		{
			public: BoundSheet8Record(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOUND_SHEET_8;
			protected: void SetDefaults();
			protected: unsigned int m_lbPlyPos;
			protected: unsigned char m_hsState;
			protected: unsigned char m_reserved;
			protected: unsigned char m_dt;
			protected: ShortXLUnicodeStringStruct* m_stName = 0;
			public: BoundSheet8Record(const char* sxName);
			public: const char* GetName();
			public: void SetStreamOffset(unsigned int nStreamOffset);
			public: virtual ~BoundSheet8Record();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BottomMarginRecord : public BiffRecord
		{
			public: BottomMarginRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOTTOM_MARGIN;
			protected: void SetDefaults();
			protected: double m_num;
			public: BottomMarginRecord(double num);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoolErrRecord : public BiffRecord
		{
			public: BoolErrRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 8;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOOLERR;
			protected: void SetDefaults();
			protected: CellStruct* m_cell = 0;
			protected: unsigned char m_bBoolErr;
			protected: unsigned char m_fError;
			public: BoolErrRecord(unsigned short nX, unsigned short nY, unsigned short nXfIndex, bool bBoolean);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: bool IsBoolean();
			public: bool GetBoolean();
			public: virtual ~BoolErrRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BookExtRecord : public BiffRecord
		{
			public: BookExtRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 20;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOOK_EXT;
			protected: void SetDefaults();
			protected: FrtHeaderStruct* m_frtHeader = 0;
			protected: unsigned int m_cb;
			protected: unsigned int m_fDontAutoRecover;
			protected: unsigned int m_fHidePivotList;
			protected: unsigned int m_fFilterPrivacy;
			protected: unsigned int m_fEmbedFactoids;
			protected: unsigned int m_mdFactoidDisplay;
			protected: unsigned int m_fSavedDuringRecovery;
			protected: unsigned int m_fCreatedViaMinimalSave;
			protected: unsigned int m_fOpenedViaDataRecovery;
			protected: unsigned int m_fOpenedViaSafeLoad;
			protected: unsigned int m_reserved;
			protected: unsigned char m_grbit1;
			protected: unsigned char m_grbit2;
			public: BookExtRecord();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~BookExtRecord();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BookBool : public BiffRecord
		{
			public: BookBool();
			public: BookBool(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOOK_BOOL;
			protected: void SetDefaults();
			protected: unsigned short m_fNoSaveSup;
			protected: unsigned short m_reserved1;
			protected: unsigned short m_fHasEnvelope;
			protected: unsigned short m_fEnvelopeVisible;
			protected: unsigned short m_fEnvelopeInitDone;
			protected: unsigned short m_grUpdateLinks;
			protected: unsigned short m_unused;
			protected: unsigned short m_fHideBorderUnselLists;
			protected: unsigned short m_reserved2;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BofRecord : public BiffRecord
		{
			public: BofRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BOF;
			protected: void SetDefaults();
			protected: unsigned short m_vers;
			protected: unsigned short m_dt;
			protected: unsigned short m_rupBuild;
			protected: unsigned short m_rupYear;
			protected: unsigned int m_nFileHistoryFlags;
			protected: unsigned int m_nMinimumExcelVersion;
			public: enum BofType
			{
				BOF_TYPE_WORKBOOK_GLOBALS = 0x005,
				BOF_TYPE_VISUAL_BASIC_MODULE = 0x0006,
				BOF_TYPE_SHEET = 0x0010,
				BOF_TYPE_CHART = 0x0020,
				BOF_TYPE_MACRO_SHEET = 0x0040,
			};

			public: BofRecord(BofType eType);
			public: BofRecord::BofType GetBofType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Blank : public BiffRecord
		{
			public: Blank(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 6;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BLANK;
			protected: void SetDefaults();
			protected: CellStruct* m_cell = 0;
			public: Blank(unsigned short nX, unsigned short nY, unsigned short nXfIndex);
			public: unsigned short GetX();
			public: unsigned short GetY();
			public: unsigned short GetXfIndex();
			public: virtual ~Blank();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffRecord_ContinueInfo
		{
			public: int m_nOffset;
			public: int m_nType;
			public: BiffRecord_ContinueInfo(int nOffset, int nType);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BeginRecord : public BiffRecord
		{
			public: BeginRecord();
			public: BeginRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 0;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Begin;
			protected: void SetDefaults();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BarRecord : public BiffRecord
		{
			public: BarRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 6;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Bar;
			protected: void SetDefaults();
			protected: short m_pcOverlap;
			protected: unsigned short m_pcGap;
			protected: unsigned short m_fTranspose;
			protected: unsigned short m_fStacked;
			protected: unsigned short m_f100;
			protected: unsigned short m_fHasShadow;
			protected: unsigned short m_reserved;
			public: BarRecord(Chart::Type eType);
			public: Chart::Type GetChartType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Backup : public BiffRecord
		{
			public: Backup();
			public: Backup(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_BACKUP;
			protected: void SetDefaults();
			protected: unsigned short m_fBackup;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisRecord : public BiffRecord
		{
			public: AxisRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Axis;
			protected: void SetDefaults();
			protected: unsigned short m_wType;
			protected: unsigned int m_reserved1;
			protected: unsigned int m_reserved2;
			protected: unsigned int m_reserved3;
			protected: unsigned int m_reserved4;
			public: AxisRecord(unsigned short wType);
			public: unsigned short GetType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisParentRecord : public BiffRecord
		{
			public: AxisParentRecord();
			public: AxisParentRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_AxisParent;
			protected: void SetDefaults();
			protected: unsigned short m_iax;
			protected: unsigned int m_unused1;
			protected: unsigned int m_unused2;
			protected: unsigned int m_unused3;
			protected: unsigned int m_unused4;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisLineRecord : public BiffRecord
		{
			public: AxisLineRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_AxisLine;
			protected: void SetDefaults();
			protected: unsigned short m_id;
			public: AxisLineRecord(unsigned short id);
			public: unsigned short GetId();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxesUsedRecord : public BiffRecord
		{
			public: AxesUsedRecord();
			public: AxesUsedRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_AxesUsed;
			protected: void SetDefaults();
			protected: unsigned short m_cAxes;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxcExtRecord : public BiffRecord
		{
			public: AxcExtRecord();
			public: AxcExtRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 18;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_AxcExt;
			protected: void SetDefaults();
			protected: unsigned short m_catMin;
			protected: unsigned short m_catMax;
			protected: unsigned short m_catMajor;
			protected: unsigned short m_duMajor;
			protected: unsigned short m_catMinor;
			protected: unsigned short m_duMinor;
			protected: unsigned short m_duBase;
			protected: unsigned short m_catCrossDate;
			protected: unsigned short m_fAutoMin;
			protected: unsigned short m_fAutoMax;
			protected: unsigned short m_fAutoMajor;
			protected: unsigned short m_fAutoMinor;
			protected: unsigned short m_fDateAxis;
			protected: unsigned short m_fAutoBase;
			protected: unsigned short m_fAutoCross;
			protected: unsigned short m_fAutoDate;
			protected: unsigned short m_reserved;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaRecord : public BiffRecord
		{
			public: AreaRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 2;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_Area;
			protected: void SetDefaults();
			protected: unsigned short m_fStacked;
			protected: unsigned short m_f100;
			protected: unsigned short m_fHasShadow;
			protected: unsigned short m_reserved;
			public: AreaRecord(Chart::Type eType);
			public: Chart::Type GetChartType();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaFormatRecord : public BiffRecord
		{
			public: AreaFormatRecord(BiffHeader* pHeader, Stream* pStream);
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			protected: static const unsigned short SIZE = 16;
			protected: static const BiffRecord::Type TYPE = BiffRecord::Type::TYPE_AreaFormat;
			protected: void SetDefaults();
			protected: unsigned int m_rgbFore;
			protected: unsigned int m_rgbBack;
			protected: unsigned short m_fls;
			protected: unsigned short m_fAuto;
			protected: unsigned short m_fInvertNeg;
			protected: unsigned short m_reserved;
			protected: unsigned short m_icvFore;
			protected: unsigned short m_icvBack;
			public: AreaFormatRecord(unsigned int rgbFore, unsigned int rgbBack, unsigned short fls, bool fAuto, bool fInvertNeg, unsigned short icvFore, unsigned short icvBack);
			public: AreaFormatRecord(Fill* pFill);
			public: void ModifyFill(Fill* pFill, BiffWorkbookGlobals* pBiffWorkbookGlobals);
			public: void PackForChecksum(BlobView* pBlobView);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XTIStruct : public BiffStruct
		{
			public: XTIStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: unsigned short m_iSupBook;
			public: short m_itabFirst;
			public: short m_itabLast;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeStringStruct : public BiffStruct
		{
			public: XLUnicodeStringStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 3;
			public: void SetDefaults();
			public: unsigned short m_cch;
			public: unsigned char m_fHighByte;
			public: unsigned char m_reserved;
			public: InternalString* m_rgb = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: int GetDynamicSize();
			public: virtual ~XLUnicodeStringStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeRichExtendedString_ContinueInfo
		{
			public: int m_nOffset;
			public: unsigned char m_fHighByte;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeRichExtendedString : public BiffStruct
		{
			public: XLUnicodeRichExtendedString();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 3;
			public: void SetDefaults();
			public: unsigned short m_cch;
			public: unsigned char m_fHighByte;
			public: unsigned char m_reserved1;
			public: unsigned char m_fExtSt;
			public: unsigned char m_fRichSt;
			public: unsigned char m_reserved2;
			public: unsigned short m_cRun;
			public: int m_cbExtRst;
			public: InternalString* m_rgb = 0;
			public: OwnedVector<XLUnicodeRichExtendedString_ContinueInfo*>* m_pContinueInfoVector = 0;
			public: XLUnicodeRichExtendedString(const char* szString);
			protected: void PostSetDefaults();
			public: void ContinueAwareBlobRead(BlobView* pBlobView, OwnedVector<BiffRecord_ContinueInfo*>* pContinueInfoVector);
			protected: void PostBlobRead(BlobView* pBlobView);
			public: void ContinueAwareBlobWrite(BlobView* pBlobView, OwnedVector<BiffRecord_ContinueInfo*>* pContinueInfoVector);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~XLUnicodeRichExtendedString();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShortXLUnicodeStringStruct : public BiffStruct
		{
			public: ShortXLUnicodeStringStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned char m_cch;
			public: unsigned char m_fHighByte;
			public: unsigned char m_reserved;
			public: InternalString* m_rgb = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: int GetDynamicSize();
			public: virtual ~ShortXLUnicodeStringStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RwUStruct : public BiffStruct
		{
			public: RwUStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_rw;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RwStruct : public BiffStruct
		{
			public: RwStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_rw;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RkRecStruct : public BiffStruct
		{
			public: RkRecStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: unsigned short m_ixfe;
			public: unsigned int m_RK;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceStruct : public BiffStruct
		{
			public: RgceStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 0;
			public: void SetDefaults();
			public: OwnedVector<ParsedExpressionRecord*>* m_pParsedExpressionRecordVector = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: int GetSize();
			public: virtual ~RgceStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceLocStruct : public BiffStruct
		{
			public: RgceLocStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: RwUStruct* m_row = 0;
			public: ColRelUStruct* m_column = 0;
			public: virtual ~RgceLocStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceAreaStruct : public BiffStruct
		{
			public: RgceAreaStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: RwUStruct* m_rowFirst = 0;
			public: RwUStruct* m_rowLast = 0;
			public: ColRelUStruct* m_columnFirst = 0;
			public: ColRelUStruct* m_columnLast = 0;
			public: virtual ~RgceAreaStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Ref8Struct : public BiffStruct
		{
			public: Ref8Struct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: unsigned short m_rwFirst;
			public: unsigned short m_rwLast;
			public: unsigned short m_colFirst;
			public: unsigned short m_colLast;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSpaceTypeStruct : public BiffStruct
		{
			public: PtgAttrSpaceTypeStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned char m_type;
			public: unsigned char m_cch;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtRecordHeaderStruct : public BiffStruct
		{
			public: OfficeArtRecordHeaderStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: unsigned short m_recVer;
			public: unsigned short m_recInstance;
			public: unsigned short m_recType;
			public: unsigned int m_recLen;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtIDCLStruct : public BiffStruct
		{
			public: OfficeArtIDCLStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: unsigned int m_dgid;
			public: unsigned int m_cspidCur;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFRITStruct : public BiffStruct
		{
			public: OfficeArtFRITStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: unsigned short m_fridNew;
			public: unsigned short m_fridOld;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTEStruct : public BiffStruct
		{
			public: OfficeArtFOPTEStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: OfficeArtFOPTEOPIDStruct* m_opid = 0;
			public: int m_op;
			public: Blob* m_pComplexData = 0;
			protected: void PostSetDefaults();
			public: virtual ~OfficeArtFOPTEStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTEOPIDStruct : public BiffStruct
		{
			public: OfficeArtFOPTEOPIDStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_opid;
			public: unsigned short m_fBid;
			public: unsigned short m_fComplex;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MSOCRStruct : public BiffStruct
		{
			public: MSOCRStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: unsigned char m_red;
			public: unsigned char m_green;
			public: unsigned char m_blue;
			public: unsigned char m_unused1;
			public: unsigned char m_fSchemeIndex;
			public: unsigned char m_unused2;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MD4DigestStruct : public BiffStruct
		{
			public: MD4DigestStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 16;
			public: void SetDefaults();
			public: unsigned char m_rgbUid1_0;
			public: unsigned char m_rgbUid1_1;
			public: unsigned char m_rgbUid1_2;
			public: unsigned char m_rgbUid1_3;
			public: unsigned char m_rgbUid1_4;
			public: unsigned char m_rgbUid1_5;
			public: unsigned char m_rgbUid1_6;
			public: unsigned char m_rgbUid1_7;
			public: unsigned char m_rgbUid1_8;
			public: unsigned char m_rgbUid1_9;
			public: unsigned char m_rgbUid1_10;
			public: unsigned char m_rgbUid1_11;
			public: unsigned char m_rgbUid1_12;
			public: unsigned char m_rgbUid1_13;
			public: unsigned char m_rgbUid1_14;
			public: unsigned char m_rgbUid1_15;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IcvFontStruct : public BiffStruct
		{
			public: IcvFontStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_icv;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IXFCellStruct : public BiffStruct
		{
			public: IXFCellStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_ixfe;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IHlinkStruct : public BiffStruct
		{
			public: IHlinkStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 24;
			public: void SetDefaults();
			public: unsigned char m_CLSID_StdHlink_0;
			public: unsigned char m_CLSID_StdHlink_1;
			public: unsigned char m_CLSID_StdHlink_2;
			public: unsigned char m_CLSID_StdHlink_3;
			public: unsigned char m_CLSID_StdHlink_4;
			public: unsigned char m_CLSID_StdHlink_5;
			public: unsigned char m_CLSID_StdHlink_6;
			public: unsigned char m_CLSID_StdHlink_7;
			public: unsigned char m_CLSID_StdHlink_8;
			public: unsigned char m_CLSID_StdHlink_9;
			public: unsigned char m_CLSID_StdHlink_10;
			public: unsigned char m_CLSID_StdHlink_11;
			public: unsigned char m_CLSID_StdHlink_12;
			public: unsigned char m_CLSID_StdHlink_13;
			public: unsigned char m_CLSID_StdHlink_14;
			public: unsigned char m_CLSID_StdHlink_15;
			public: HyperlinkObjectStruct* m_hyperlink = 0;
			public: virtual ~IHlinkStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HyperlinkStringStruct : public BiffStruct
		{
			public: HyperlinkStringStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: unsigned int m_length;
			protected: InternalString* m_string = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~HyperlinkStringStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HyperlinkObjectStruct : public BiffStruct
		{
			public: HyperlinkObjectStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: unsigned int m_streamVersion;
			public: unsigned int m_hlstmfHasMoniker;
			public: unsigned int m_hlstmfIsAbsolute;
			public: unsigned int m_hlstmfSiteGaveDisplayName;
			public: unsigned int m_hlstmfHasLocationStr;
			public: unsigned int m_hlstmfHasDisplayName;
			public: unsigned int m_hlstmfHasGUID;
			public: unsigned int m_hlstmfHasCreationTime;
			public: unsigned int m_hlstmfHasFrameName;
			public: unsigned int m_hlstmfMonikerSavedAsStr;
			public: unsigned int m_hlstmfAbsFromGetdataRel;
			public: unsigned int m_reserved;
			protected: HyperlinkStringStruct* m_displayName = 0;
			protected: HyperlinkStringStruct* m_targetFrameName = 0;
			protected: HyperlinkStringStruct* m_moniker = 0;
			public: InternalString* m_haxUrl = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~HyperlinkObjectStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FullColorExtStruct : public BiffStruct
		{
			public: FullColorExtStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 16;
			public: void SetDefaults();
			public: unsigned short m_xclrType;
			public: short m_nTintShade;
			public: unsigned int m_xclrValue;
			public: unsigned char m_unusedA;
			public: unsigned char m_unusedB;
			public: unsigned char m_unusedC;
			public: unsigned char m_unusedD;
			public: unsigned char m_unusedE;
			public: unsigned char m_unusedF;
			public: unsigned char m_unusedG;
			public: unsigned char m_unusedH;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtPioGrbitStruct : public BiffStruct
		{
			public: FtPioGrbitStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: unsigned short m_ft;
			public: unsigned short m_cb;
			public: unsigned short m_fAutoPict;
			public: unsigned short m_fDde;
			public: unsigned short m_fPrintCalc;
			public: unsigned short m_fIcon;
			public: unsigned short m_fCtl;
			public: unsigned short m_fPrstm;
			public: unsigned short m_unused1;
			public: unsigned short m_fCamera;
			public: unsigned short m_fDefaultSize;
			public: unsigned short m_fAutoLoad;
			public: unsigned short m_unused2;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtCfStruct : public BiffStruct
		{
			public: FtCfStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: unsigned short m_ft;
			public: unsigned short m_cb;
			public: unsigned short m_cf;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtHeaderStruct : public BiffStruct
		{
			public: FrtHeaderStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 12;
			public: void SetDefaults();
			public: unsigned short m_rt;
			public: FrtFlagsStruct* m_grbitFrt = 0;
			public: unsigned int m_reservedA;
			public: unsigned int m_reservedB;
			public: virtual ~FrtHeaderStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtHeaderOldStruct : public BiffStruct
		{
			public: FrtHeaderOldStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: unsigned short m_rt;
			public: FrtFlagsStruct* m_grbitFrt = 0;
			public: virtual ~FrtHeaderOldStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtFlagsStruct : public BiffStruct
		{
			public: FrtFlagsStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_fFrtRef;
			public: unsigned short m_fFrtAlert;
			public: unsigned short m_reserved;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FormulaValueStruct : public BiffStruct
		{
			public: FormulaValueStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 8;
			public: void SetDefaults();
			public: unsigned char m_byte1;
			public: unsigned char m_byte2;
			public: unsigned char m_byte3;
			public: unsigned char m_byte4;
			public: unsigned char m_byte5;
			public: unsigned char m_byte6;
			public: unsigned short m_fExprO;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ExtPropStruct : public BiffStruct
		{
			public: ExtPropStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 4;
			public: void SetDefaults();
			public: unsigned short m_extType;
			public: unsigned short m_cb;
			public: FullColorExtStruct* m_pFullColorExt = 0;
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: virtual ~ExtPropStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColStruct : public BiffStruct
		{
			public: ColStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_col;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColRelUStruct : public BiffStruct
		{
			public: ColRelUStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_col;
			public: unsigned short m_colRelative;
			public: unsigned short m_rowRelative;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellStruct : public BiffStruct
		{
			public: CellStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 6;
			public: void SetDefaults();
			public: RwStruct* m_rw = 0;
			public: ColStruct* m_col = 0;
			public: IXFCellStruct* m_ixfe = 0;
			public: virtual ~CellStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellParsedFormulaStruct : public BiffStruct
		{
			public: CellParsedFormulaStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned short m_cce;
			public: RgceStruct* m_rgce = 0;
			public: CellParsedFormulaStruct(Formula* pFormula, WorkbookGlobals* pWorkbookGlobals);
			protected: void PostSetDefaults();
			protected: void PostBlobRead(BlobView* pBlobView);
			protected: void PreBlobWrite(BlobView* pBlobView);
			protected: void PostBlobWrite(BlobView* pBlobView);
			public: int GetSize();
			public: virtual ~CellParsedFormulaStruct();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BuiltInStyleStruct : public BiffStruct
		{
			public: BuiltInStyleStruct();
			public: virtual void BlobRead(BlobView* pBlobView);
			public: virtual void BlobWrite(BlobView* pBlobView);
			public: static const unsigned short SIZE = 2;
			public: void SetDefaults();
			public: unsigned char m_istyBuiltIn;
			public: unsigned char m_iLevel;
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDimensions
		{
			public: unsigned short m_nCellX1;
			public: unsigned short m_nSubCellX1;
			public: unsigned short m_nCellY1;
			public: unsigned short m_nSubCellY1;
			public: unsigned short m_nCellX2;
			public: unsigned short m_nSubCellX2;
			public: unsigned short m_nCellY2;
			public: unsigned short m_nSubCellY2;
		};
		class BiffWorksheet : public Worksheet
		{
			protected: BiffRecordContainer* m_pBiffRecordContainer = 0;
			public: BiffWorksheet(Workbook* pWorkbook);
			public: BiffWorksheet(Workbook* pWorkbook, BiffWorkbookGlobals* pBiffWorkbookGlobals, BiffRecord* pInitialBiffRecord, Stream* pStream);
			public: static void Write(Worksheet* pWorksheet, WorkbookGlobals* pWorkbookGlobals, unsigned short nWorksheetIndex, BiffRecordContainer* pBiffRecordContainer);
			public: static OfficeArtDimensions* ComputeDimensions(Worksheet* pWorksheet, unsigned short nCellX, unsigned short nSubCellX, unsigned short nWidth, unsigned short nCellY, unsigned short nSubCellY, unsigned short nHeight);
			public: int LoopToEnd(int i, BiffRecordContainer* pBiffRecordContainer);
			public: virtual ~BiffWorksheet();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffUtils
		{
			public: static double RkValueDecode(unsigned int nRkValue);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffRecordContainer : public WorkbookGlobals
		{
			public: OwnedVector<BiffRecord*>* m_pBiffRecordVector = 0;
			public: BiffRecordContainer();
			public: BiffRecordContainer(BiffRecord* pInitialBiffRecord, Stream* pStream);
			public: void AddBiffRecord(BiffRecord* pBiffRecord);
			public: unsigned int GetSize();
			public: void Write(Stream* pStream);
			public: virtual ~BiffRecordContainer();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDirectoryImplementation
		{
			public: unsigned int m_nMinimumStandardStreamSize;
			public: OwnedVector<Stream*>* m_pStreamVector = 0;
			public: CompoundFile* m_pCompoundFile = 0;
			public: void AppendStream(Stream* pStream);
			public: static int RedBlackTreeComparisonCallback(const void* pObjectA, const void* pObjectB);
			public: void RedBlackTreeWalk(RedBlackNode* pNode, Stream* pStorage);
			public: virtual ~StreamDirectoryImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDirectory : public SectorChain
		{
			protected: StreamDirectoryImplementation* m_pImpl = 0;
			public: StreamDirectory(int nSectorSize, unsigned int nMinimumStandardStreamSize, CompoundFile* pCompoundFile);
			public: int GetNumStream();
			public: Stream* GetStreamByIndex(int nStreamDirectoryId);
			public: Stream* GetStreamByName(const char* sxName);
			public: Stream* CreateStream(const char* sxName, Stream::Type eType);
			public: virtual void AppendSector(Sector* pSector);
			public: virtual void Extend(Sector* pSector);
			public: void RedBlackTreeRebuild();
			public: virtual ~StreamDirectory();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorAllocationTable : public SectorChain
		{
			protected: int m_nNumFreeSectorId;
			public: SectorAllocationTable(int nSectorSize);
			public: virtual int GetNumSectorId();
			public: virtual int GetNumFreeSectorId();
			public: virtual int GetSectorId(int nIndex);
			public: virtual void SetSectorId(int nIndex, int nSectorId);
			public: int GetFreeSectorId();
			public: virtual void AppendSector(Sector* pSector);
			public: virtual void Extend(Sector* pSector);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorImplementation
		{
			public: int m_nSectorId;
			public: BlobView* m_pBlobView = 0;
			public: virtual ~SectorImplementation();
		};
		class Sector
		{
			public: static const int MINIMUM_SECTOR_SIZE = 128;
			public: static const int MINIMUM_SECTOR_SIZE_SHIFT = 7;
			protected: SectorImplementation* m_pImpl = 0;
			public: enum SectorId
			{
				FREE_SECTOR_SECTOR_ID = -1,
				END_OF_CHAIN_SECTOR_ID = -2,
				SECTOR_ALLOCATION_TABLE_SECTOR_ID = -3,
				MASTER_SECTOR_ALLOCATION_TABLE_SECTOR_ID = -4,
			};

			public: Sector(int nSectorId, BlobView* pBlobView, int nDataSize);
			public: int GetDataSize();
			public: int GetSectorId();
			public: BlobView* GetBlobView();
			public: virtual ~Sector();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MasterSectorAllocationTable : public SectorAllocationTable
		{
			public: static const int INITIAL_SECTOR_ID_ARRAY_SIZE = 109;
			protected: CompoundFileHeader* m_pHeader = 0;
			public: MasterSectorAllocationTable(CompoundFileHeader* pHeader);
			public: virtual int GetNumSectorId();
			public: int GetNumInternalSectorId();
			public: virtual int GetSectorId(int nIndex);
			public: virtual void SetSectorId(int nIndex, int nSectorId);
			public: int GetInternalSectorId(int nIndex);
			public: void SetInternalSectorId(int nIndex, int nSectorId);
			public: int TranslateIndex(int nIndex);
			public: int GetSectorIdToAppend();
			public: virtual void AppendSector(Sector* pSector);
			public: virtual void Extend(Sector* pSector);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CompoundFileHeader
		{
			public: static const int MAGIC_WORD_SIZE = 8;
			public: static const unsigned char MAGIC_WORD[CompoundFileHeader::MAGIC_WORD_SIZE];
			public: static const unsigned int UNIQUE_IDENTIFIER_SIZE = 16;
			public: static const unsigned short REVISION_NUMBER = 62;
			public: static const unsigned short VERSION_NUMBER = 0x0003;
			public: static const unsigned short BOM_LITTLE_ENDIAN = 0xFFFF;
			public: static const unsigned short BOM_BIG_ENDIAN = 0xFFFE;
			public: unsigned char m_pMagicWord[MAGIC_WORD_SIZE];
			public: unsigned char m_pUniqueIdentifier[UNIQUE_IDENTIFIER_SIZE];
			public: unsigned short m_nRevisonNumber;
			public: unsigned short m_nVersionNumber;
			public: unsigned short m_nByteOrderIdentifier;
			public: unsigned short m_nSectorSize;
			public: unsigned short m_nShortSectorSize;
			public: unsigned char m_pUnusedA[10];
			public: unsigned int m_nSectorAllocationTableSize;
			public: int m_nStreamDirectoryStreamSectorId;
			public: unsigned char m_pUnusedB[4];
			public: unsigned int m_nMinimumStandardStreamSize;
			public: int m_nShortSectorAllocationTableSectorId;
			public: unsigned int m_nShortSectorAllocationTableSize;
			public: int m_nMasterSectorAllocationTableSectorId;
			public: unsigned int m_nMasterSectorAllocationTableSize;
			public: int m_pMasterSectorAllocationTable[MasterSectorAllocationTable::INITIAL_SECTOR_ID_ARRAY_SIZE];
			public: CompoundFileHeader(unsigned int nSectorSize, unsigned int nShortSectorSize);
			public: int GetSectorSize();
			public: bool Unpack(Blob* pBlob);
			public: void Pack(Blob* pBlob);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CompoundFile
		{
			public: static const int HEADER_SIZE = 512;
			protected: CompoundFileHeader* m_pHeader = 0;
			protected: Blob* m_pBlob = 0;
			protected: MasterSectorAllocationTable* m_pMasterSectorAllocationTable = 0;
			protected: SectorAllocationTable* m_pSectorAllocationTable = 0;
			protected: SectorAllocationTable* m_pShortSectorAllocationTable = 0;
			protected: StreamDirectory* m_pStreamDirectory = 0;
			protected: OwnedVector<Sector*>* m_pSectorVector = 0;
			protected: OwnedVector<Sector*>* m_pShortSectorVector = 0;
			public: CompoundFile(unsigned int nSectorSize = 512, unsigned int nShortSectorSize = 64, unsigned int nMinimumStandardStreamSize = 4096);
			public: virtual ~CompoundFile();
			public: bool Load(const char* sxFileName);
			public: bool Save(const char* sxFileName);
			public: CompoundFileHeader* GetHeader();
			public: StreamDirectory* GetStreamDirectory();
			public: Stream* CreateStream(const char* sxName, Stream::Type eType);
			public: int GetNumStream();
			public: Stream* GetStreamByIndex(int nIndex);
			public: Stream* GetStreamByName(const char* sxName);
			public: int GetSectorSize(bool bShortSector);
			public: int GetSectorId(int nSectorId, bool bShortSector);
			public: Sector* GetSector(int nSectorId, bool bShortSector);
			public: void FillSectorChain(SectorChain* pSectorChain, int nInitialSectorId, bool bShortSector);
			public: int GetFreeSectorId(bool bShortSector);
			public: SectorAllocationTable* GetSectorAllocationTable(bool bShortSector);
			public: MasterSectorAllocationTable* GetMasterSectorAllocationTable();
			public: void MasterSectorAllocationTableExtend();
			public: void SectorAllocationTableExtend(bool bShortSector);
			public: void SectorChainExtend(SectorChain* pSectorChain);
			public: void StreamDirectoryExtend();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Token
		{
			public: enum Type
			{
				TYPE_FUNC_COUNT = 0x0000,
				TYPE_FUNC_IF = 0x0001,
				TYPE_FUNC_ISNA = 0x0002,
				TYPE_FUNC_ISERROR = 0x0003,
				TYPE_FUNC_SUM = 0x0004,
				TYPE_FUNC_AVERAGE = 0x0005,
				TYPE_FUNC_MIN = 0x0006,
				TYPE_FUNC_MAX = 0x0007,
				TYPE_FUNC_ROW = 0x0008,
				TYPE_FUNC_COLUMN = 0x0009,
				TYPE_FUNC_NA = 0x000A,
				TYPE_FUNC_NPV = 0x000B,
				TYPE_FUNC_STDEV = 0x000C,
				TYPE_FUNC_DOLLAR = 0x000D,
				TYPE_FUNC_FIXED = 0x000E,
				TYPE_FUNC_SIN = 0x000F,
				TYPE_FUNC_COS = 0x0010,
				TYPE_FUNC_TAN = 0x0011,
				TYPE_FUNC_ATAN = 0x0012,
				TYPE_FUNC_PI = 0x0013,
				TYPE_FUNC_SQRT = 0x0014,
				TYPE_FUNC_EXP = 0x0015,
				TYPE_FUNC_LN = 0x0016,
				TYPE_FUNC_LOG10 = 0x0017,
				TYPE_FUNC_ABS = 0x0018,
				TYPE_FUNC_INT = 0x0019,
				TYPE_FUNC_SIGN = 0x001A,
				TYPE_FUNC_ROUND = 0x001B,
				TYPE_FUNC_LOOKUP = 0x001C,
				TYPE_FUNC_INDEX = 0x001D,
				TYPE_FUNC_REPT = 0x001E,
				TYPE_FUNC_MID = 0x001F,
				TYPE_FUNC_LEN = 0x0020,
				TYPE_FUNC_VALUE = 0x0021,
				TYPE_FUNC_TRUE = 0x0022,
				TYPE_FUNC_FALSE = 0x0023,
				TYPE_FUNC_AND = 0x0024,
				TYPE_FUNC_OR = 0x0025,
				TYPE_FUNC_NOT = 0x0026,
				TYPE_FUNC_MOD = 0x0027,
				TYPE_FUNC_DCOUNT = 0x0028,
				TYPE_FUNC_DSUM = 0x0029,
				TYPE_FUNC_DAVERAGE = 0x002A,
				TYPE_FUNC_DMIN = 0x002B,
				TYPE_FUNC_DMAX = 0x002C,
				TYPE_FUNC_DSTDEV = 0x002D,
				TYPE_FUNC_VAR = 0x002E,
				TYPE_FUNC_DVAR = 0x002F,
				TYPE_FUNC_TEXT = 0x0030,
				TYPE_FUNC_LINEST = 0x0031,
				TYPE_FUNC_TREND = 0x0032,
				TYPE_FUNC_LOGEST = 0x0033,
				TYPE_FUNC_GROWTH = 0x0034,
				TYPE_FUNC_GOTO = 0x0035,
				TYPE_FUNC_HALT = 0x0036,
				TYPE_FUNC_RETURN = 0x0037,
				TYPE_FUNC_PV = 0x0038,
				TYPE_FUNC_FV = 0x0039,
				TYPE_FUNC_NPER = 0x003A,
				TYPE_FUNC_PMT = 0x003B,
				TYPE_FUNC_RATE = 0x003C,
				TYPE_FUNC_MIRR = 0x003D,
				TYPE_FUNC_IRR = 0x003E,
				TYPE_FUNC_RAND = 0x003F,
				TYPE_FUNC_MATCH = 0x0040,
				TYPE_FUNC_DATE = 0x0041,
				TYPE_FUNC_TIME = 0x0042,
				TYPE_FUNC_DAY = 0x0043,
				TYPE_FUNC_MONTH = 0x0044,
				TYPE_FUNC_YEAR = 0x0045,
				TYPE_FUNC_WEEKDAY = 0x0046,
				TYPE_FUNC_HOUR = 0x0047,
				TYPE_FUNC_MINUTE = 0x0048,
				TYPE_FUNC_SECOND = 0x0049,
				TYPE_FUNC_NOW = 0x004A,
				TYPE_FUNC_AREAS = 0x004B,
				TYPE_FUNC_ROWS = 0x004C,
				TYPE_FUNC_COLUMNS = 0x004D,
				TYPE_FUNC_OFFSET = 0x004E,
				TYPE_FUNC_ABSREF = 0x004F,
				TYPE_FUNC_RELREF = 0x0050,
				TYPE_FUNC_ARGUMENT = 0x0051,
				TYPE_FUNC_SEARCH = 0x0052,
				TYPE_FUNC_TRANSPOSE = 0x0053,
				TYPE_FUNC_ERROR = 0x0054,
				TYPE_FUNC_STEP = 0x0055,
				TYPE_FUNC_TYPE = 0x0056,
				TYPE_FUNC_ECHO = 0x0057,
				TYPE_FUNC_SET_NAME = 0x0058,
				TYPE_FUNC_CALLER = 0x0059,
				TYPE_FUNC_DEREF = 0x005A,
				TYPE_FUNC_WINDOWS = 0x005B,
				TYPE_FUNC_SERIES = 0x005C,
				TYPE_FUNC_DOCUMENTS = 0x005D,
				TYPE_FUNC_ACTIVE_CELL = 0x005E,
				TYPE_FUNC_SELECTION = 0x005F,
				TYPE_FUNC_RESULT = 0x0060,
				TYPE_FUNC_ATAN2 = 0x0061,
				TYPE_FUNC_ASIN = 0x0062,
				TYPE_FUNC_ACOS = 0x0063,
				TYPE_FUNC_CHOOSE = 0x0064,
				TYPE_FUNC_HLOOKUP = 0x0065,
				TYPE_FUNC_VLOOKUP = 0x0066,
				TYPE_FUNC_LINKS = 0x0067,
				TYPE_FUNC_INPUT = 0x0068,
				TYPE_FUNC_ISREF = 0x0069,
				TYPE_FUNC_GET_FORMULA = 0x006A,
				TYPE_FUNC_GET_NAME = 0x006B,
				TYPE_FUNC_SET_VALUE = 0x006C,
				TYPE_FUNC_LOG = 0x006D,
				TYPE_FUNC_EXEC = 0x006E,
				TYPE_FUNC_CHAR = 0x006F,
				TYPE_FUNC_LOWER = 0x0070,
				TYPE_FUNC_UPPER = 0x0071,
				TYPE_FUNC_PROPER = 0x0072,
				TYPE_FUNC_LEFT = 0x0073,
				TYPE_FUNC_RIGHT = 0x0074,
				TYPE_FUNC_EXACT = 0x0075,
				TYPE_FUNC_TRIM = 0x0076,
				TYPE_FUNC_REPLACE = 0x0077,
				TYPE_FUNC_SUBSTITUTE = 0x0078,
				TYPE_FUNC_CODE = 0x0079,
				TYPE_FUNC_NAMES = 0x007A,
				TYPE_FUNC_DIRECTORY = 0x007B,
				TYPE_FUNC_FIND = 0x007C,
				TYPE_FUNC_CELL = 0x007D,
				TYPE_FUNC_ISERR = 0x007E,
				TYPE_FUNC_ISTEXT = 0x007F,
				TYPE_FUNC_ISNUMBER = 0x0080,
				TYPE_FUNC_ISBLANK = 0x0081,
				TYPE_FUNC_T = 0x0082,
				TYPE_FUNC_N = 0x0083,
				TYPE_FUNC_FOPEN = 0x0084,
				TYPE_FUNC_FCLOSE = 0x0085,
				TYPE_FUNC_FSIZE = 0x0086,
				TYPE_FUNC_FREADLN = 0x0087,
				TYPE_FUNC_FREAD = 0x0088,
				TYPE_FUNC_FWRITELN = 0x0089,
				TYPE_FUNC_FWRITE = 0x008A,
				TYPE_FUNC_FPOS = 0x008B,
				TYPE_FUNC_DATEVALUE = 0x008C,
				TYPE_FUNC_TIMEVALUE = 0x008D,
				TYPE_FUNC_SLN = 0x008E,
				TYPE_FUNC_SYD = 0x008F,
				TYPE_FUNC_DDB = 0x0090,
				TYPE_FUNC_GET_DEF = 0x0091,
				TYPE_FUNC_REFTEXT = 0x0092,
				TYPE_FUNC_TEXTREF = 0x0093,
				TYPE_FUNC_INDIRECT = 0x0094,
				TYPE_FUNC_REGISTER = 0x0095,
				TYPE_FUNC_CALL = 0x0096,
				TYPE_FUNC_ADD_BAR = 0x0097,
				TYPE_FUNC_ADD_MENU = 0x0098,
				TYPE_FUNC_ADD_COMMAND = 0x0099,
				TYPE_FUNC_ENABLE_COMMAND = 0x009A,
				TYPE_FUNC_CHECK_COMMAND = 0x009B,
				TYPE_FUNC_RENAME_COMMAND = 0x009C,
				TYPE_FUNC_SHOW_BAR = 0x009D,
				TYPE_FUNC_DELETE_MENU = 0x009E,
				TYPE_FUNC_DELETE_COMMAND = 0x009F,
				TYPE_FUNC_GET_CHART_ITEM = 0x00A0,
				TYPE_FUNC_DIALOG_BOX = 0x00A1,
				TYPE_FUNC_CLEAN = 0x00A2,
				TYPE_FUNC_MDETERM = 0x00A3,
				TYPE_FUNC_MINVERSE = 0x00A4,
				TYPE_FUNC_MMULT = 0x00A5,
				TYPE_FUNC_FILES = 0x00A6,
				TYPE_FUNC_IPMT = 0x00A7,
				TYPE_FUNC_PPMT = 0x00A8,
				TYPE_FUNC_COUNTA = 0x00A9,
				TYPE_FUNC_CANCEL_KEY = 0x00AA,
				TYPE_FUNC_FOR = 0x00AB,
				TYPE_FUNC_WHILE = 0x00AC,
				TYPE_FUNC_BREAK = 0x00AD,
				TYPE_FUNC_NEXT = 0x00AE,
				TYPE_FUNC_INITIATE = 0x00AF,
				TYPE_FUNC_REQUEST = 0x00B0,
				TYPE_FUNC_POKE = 0x00B1,
				TYPE_FUNC_EXECUTE = 0x00B2,
				TYPE_FUNC_TERMINATE = 0x00B3,
				TYPE_FUNC_RESTART = 0x00B4,
				TYPE_FUNC_HELP = 0x00B5,
				TYPE_FUNC_GET_BAR = 0x00B6,
				TYPE_FUNC_PRODUCT = 0x00B7,
				TYPE_FUNC_FACT = 0x00B8,
				TYPE_FUNC_GET_CELL = 0x00B9,
				TYPE_FUNC_GET_WORKSPACE = 0x00BA,
				TYPE_FUNC_GET_WINDOW = 0x00BB,
				TYPE_FUNC_GET_DOCUMENT = 0x00BC,
				TYPE_FUNC_DPRODUCT = 0x00BD,
				TYPE_FUNC_ISNONTEXT = 0x00BE,
				TYPE_FUNC_GET_NOTE = 0x00BF,
				TYPE_FUNC_NOTE = 0x00C0,
				TYPE_FUNC_STDEVP = 0x00C1,
				TYPE_FUNC_VARP = 0x00C2,
				TYPE_FUNC_DSTDEVP = 0x00C3,
				TYPE_FUNC_DVARP = 0x00C4,
				TYPE_FUNC_TRUNC = 0x00C5,
				TYPE_FUNC_ISLOGICAL = 0x00C6,
				TYPE_FUNC_DCOUNTA = 0x00C7,
				TYPE_FUNC_DELETE_BAR = 0x00C8,
				TYPE_FUNC_UNREGISTER = 0x00C9,
				TYPE_FUNC_USDOLLAR = 0x00CC,
				TYPE_FUNC_FINDB = 0x00CD,
				TYPE_FUNC_SEARCHB = 0x00CE,
				TYPE_FUNC_REPLACEB = 0x00CF,
				TYPE_FUNC_LEFTB = 0x00D0,
				TYPE_FUNC_RIGHTB = 0x00D1,
				TYPE_FUNC_MIDB = 0x00D2,
				TYPE_FUNC_LENB = 0x00D3,
				TYPE_FUNC_ROUNDUP = 0x00D4,
				TYPE_FUNC_ROUNDDOWN = 0x00D5,
				TYPE_FUNC_ASC = 0x00D6,
				TYPE_FUNC_DBCS = 0x00D7,
				TYPE_FUNC_RANK = 0x00D8,
				TYPE_FUNC_ADDRESS = 0x00DB,
				TYPE_FUNC_DAYS360 = 0x00DC,
				TYPE_FUNC_TODAY = 0x00DD,
				TYPE_FUNC_VDB = 0x00DE,
				TYPE_FUNC_ELSE = 0x00DF,
				TYPE_FUNC_ELSE_IF = 0x00E0,
				TYPE_FUNC_END_IF = 0x00E1,
				TYPE_FUNC_FOR_CELL = 0x00E2,
				TYPE_FUNC_MEDIAN = 0x00E3,
				TYPE_FUNC_SUMPRODUCT = 0x00E4,
				TYPE_FUNC_SINH = 0x00E5,
				TYPE_FUNC_COSH = 0x00E6,
				TYPE_FUNC_TANH = 0x00E7,
				TYPE_FUNC_ASINH = 0x00E8,
				TYPE_FUNC_ACOSH = 0x00E9,
				TYPE_FUNC_ATANH = 0x00EA,
				TYPE_FUNC_DGET = 0x00EB,
				TYPE_FUNC_CREATE_OBJECT = 0x00EC,
				TYPE_FUNC_VOLATILE = 0x00ED,
				TYPE_FUNC_LAST_ERROR = 0x00EE,
				TYPE_FUNC_CUSTOM_UNDO = 0x00EF,
				TYPE_FUNC_CUSTOM_REPEAT = 0x00F0,
				TYPE_FUNC_FORMULA_CONVERT = 0x00F1,
				TYPE_FUNC_GET_LINK_INFO = 0x00F2,
				TYPE_FUNC_TEXT_BOX = 0x00F3,
				TYPE_FUNC_INFO = 0x00F4,
				TYPE_FUNC_GROUP = 0x00F5,
				TYPE_FUNC_GET_OBJECT = 0x00F6,
				TYPE_FUNC_DB = 0x00F7,
				TYPE_FUNC_PAUSE = 0x00F8,
				TYPE_FUNC_RESUME = 0x00FB,
				TYPE_FUNC_FREQUENCY = 0x00FC,
				TYPE_FUNC_ADD_TOOLBAR = 0x00FD,
				TYPE_FUNC_DELETE_TOOLBAR = 0x00FE,
				TYPE_FUNC_RESET_TOOLBAR = 0x0100,
				TYPE_FUNC_EVALUATE = 0x0101,
				TYPE_FUNC_GET_TOOLBAR = 0x0102,
				TYPE_FUNC_GET_TOOL = 0x0103,
				TYPE_FUNC_SPELLING_CHECK = 0x0104,
				TYPE_FUNC_ERROR_TYPE = 0x0105,
				TYPE_FUNC_APP_TITLE = 0x0106,
				TYPE_FUNC_WINDOW_TITLE = 0x0107,
				TYPE_FUNC_SAVE_TOOLBAR = 0x0108,
				TYPE_FUNC_ENABLE_TOOL = 0x0109,
				TYPE_FUNC_PRESS_TOOL = 0x010A,
				TYPE_FUNC_REGISTER_ID = 0x010B,
				TYPE_FUNC_GET_WORKBOOK = 0x010C,
				TYPE_FUNC_AVEDEV = 0x010D,
				TYPE_FUNC_BETADIST = 0x010E,
				TYPE_FUNC_GAMMALN = 0x010F,
				TYPE_FUNC_BETAINV = 0x0110,
				TYPE_FUNC_BINOMDIST = 0x0111,
				TYPE_FUNC_CHIDIST = 0x0112,
				TYPE_FUNC_CHIINV = 0x0113,
				TYPE_FUNC_COMBIN = 0x0114,
				TYPE_FUNC_CONFIDENCE = 0x0115,
				TYPE_FUNC_CRITBINOM = 0x0116,
				TYPE_FUNC_EVEN = 0x0117,
				TYPE_FUNC_EXPONDIST = 0x0118,
				TYPE_FUNC_FDIST = 0x0119,
				TYPE_FUNC_FINV = 0x011A,
				TYPE_FUNC_FISHER = 0x011B,
				TYPE_FUNC_FISHERINV = 0x011C,
				TYPE_FUNC_FLOOR = 0x011D,
				TYPE_FUNC_GAMMADIST = 0x011E,
				TYPE_FUNC_GAMMAINV = 0x011F,
				TYPE_FUNC_CEILING = 0x0120,
				TYPE_FUNC_HYPGEOMDIST = 0x0121,
				TYPE_FUNC_LOGNORMDIST = 0x0122,
				TYPE_FUNC_LOGINV = 0x0123,
				TYPE_FUNC_NEGBINOMDIST = 0x0124,
				TYPE_FUNC_NORMDIST = 0x0125,
				TYPE_FUNC_NORMSDIST = 0x0126,
				TYPE_FUNC_NORMINV = 0x0127,
				TYPE_FUNC_NORMSINV = 0x0128,
				TYPE_FUNC_STANDARDIZE = 0x0129,
				TYPE_FUNC_ODD = 0x012A,
				TYPE_FUNC_PERMUT = 0x012B,
				TYPE_FUNC_POISSON = 0x012C,
				TYPE_FUNC_TDIST = 0x012D,
				TYPE_FUNC_WEIBULL = 0x012E,
				TYPE_FUNC_SUMXMY2 = 0x012F,
				TYPE_FUNC_SUMX2MY2 = 0x0130,
				TYPE_FUNC_SUMX2PY2 = 0x0131,
				TYPE_FUNC_CHITEST = 0x0132,
				TYPE_FUNC_CORREL = 0x0133,
				TYPE_FUNC_COVAR = 0x0134,
				TYPE_FUNC_FORECAST = 0x0135,
				TYPE_FUNC_FTEST = 0x0136,
				TYPE_FUNC_INTERCEPT = 0x0137,
				TYPE_FUNC_PEARSON = 0x0138,
				TYPE_FUNC_RSQ = 0x0139,
				TYPE_FUNC_STEYX = 0x013A,
				TYPE_FUNC_SLOPE = 0x013B,
				TYPE_FUNC_TTEST = 0x013C,
				TYPE_FUNC_PROB = 0x013D,
				TYPE_FUNC_DEVSQ = 0x013E,
				TYPE_FUNC_GEOMEAN = 0x013F,
				TYPE_FUNC_HARMEAN = 0x0140,
				TYPE_FUNC_SUMSQ = 0x0141,
				TYPE_FUNC_KURT = 0x0142,
				TYPE_FUNC_SKEW = 0x0143,
				TYPE_FUNC_ZTEST = 0x0144,
				TYPE_FUNC_LARGE = 0x0145,
				TYPE_FUNC_SMALL = 0x0146,
				TYPE_FUNC_QUARTILE = 0x0147,
				TYPE_FUNC_PERCENTILE = 0x0148,
				TYPE_FUNC_PERCENTRANK = 0x0149,
				TYPE_FUNC_MODE = 0x014A,
				TYPE_FUNC_TRIMMEAN = 0x014B,
				TYPE_FUNC_TINV = 0x014C,
				TYPE_FUNC_MOVIE_COMMAND = 0x014E,
				TYPE_FUNC_GET_MOVIE = 0x014F,
				TYPE_FUNC_CONCATENATE = 0x0150,
				TYPE_FUNC_POWER = 0x0151,
				TYPE_FUNC_PIVOT_ADD_DATA = 0x0152,
				TYPE_FUNC_GET_PIVOT_TABLE = 0x0153,
				TYPE_FUNC_GET_PIVOT_FIELD = 0x0154,
				TYPE_FUNC_GET_PIVOT_ITEM = 0x0155,
				TYPE_FUNC_RADIANS = 0x0156,
				TYPE_FUNC_DEGREES = 0x0157,
				TYPE_FUNC_SUBTOTAL = 0x0158,
				TYPE_FUNC_SUMIF = 0x0159,
				TYPE_FUNC_COUNTIF = 0x015A,
				TYPE_FUNC_COUNTBLANK = 0x015B,
				TYPE_FUNC_SCENARIO_GET = 0x015C,
				TYPE_FUNC_OPTIONS_LISTS_GET = 0x015D,
				TYPE_FUNC_ISPMT = 0x015E,
				TYPE_FUNC_DATEDIF = 0x015F,
				TYPE_FUNC_DATESTRING = 0x0160,
				TYPE_FUNC_NUMBERSTRING = 0x0161,
				TYPE_FUNC_ROMAN = 0x0162,
				TYPE_FUNC_OPEN_DIALOG = 0x0163,
				TYPE_FUNC_SAVE_DIALOG = 0x0164,
				TYPE_FUNC_VIEW_GET = 0x0165,
				TYPE_FUNC_GETPIVOTDATA = 0x0166,
				TYPE_FUNC_HYPERLINK = 0x0167,
				TYPE_FUNC_PHONETIC = 0x0168,
				TYPE_FUNC_AVERAGEA = 0x0169,
				TYPE_FUNC_MAXA = 0x016A,
				TYPE_FUNC_MINA = 0x016B,
				TYPE_FUNC_STDEVPA = 0x016C,
				TYPE_FUNC_VARPA = 0x016D,
				TYPE_FUNC_STDEVA = 0x016E,
				TYPE_FUNC_VARA = 0x016F,
				TYPE_FUNC_BAHTTEXT = 0x0170,
				TYPE_FUNC_THAIDAYOFWEEK = 0x0171,
				TYPE_FUNC_THAIDIGIT = 0x0172,
				TYPE_FUNC_THAIMONTHOFYEAR = 0x0173,
				TYPE_FUNC_THAINUMSOUND = 0x0174,
				TYPE_FUNC_THAINUMSTRING = 0x0175,
				TYPE_FUNC_THAISTRINGLENGTH = 0x0176,
				TYPE_FUNC_ISTHAIDIGIT = 0x0177,
				TYPE_FUNC_ROUNDBAHTDOWN = 0x0178,
				TYPE_FUNC_ROUNDBAHTUP = 0x0179,
				TYPE_FUNC_THAIYEAR = 0x017A,
				TYPE_FUNC_RTD = 0x017B,
				TYPE_BOOL,
				TYPE_INT,
				TYPE_NUMBER,
				TYPE_SPACE,
				TYPE_MISS_ARG,
				TYPE_STRING,
				TYPE_COORDINATE,
				TYPE_COORDINATE_3D,
				TYPE_AREA,
				TYPE_AREA_3D,
				TYPE_OPERATOR,
				TYPE_PAREN,
			};

			public: enum SubType
			{
				SUB_TYPE_FUNCTION,
				SUB_TYPE_OPERATOR,
				SUB_TYPE_VARIABLE,
			};

			protected: Type m_eType;
			protected: SubType m_eSubType;
			protected: unsigned char m_nParameterCount;
			public: Token(Type eType, SubType eSubType, unsigned char nParameterCount);
			public: virtual ~Token();
			public: Token::Type GetType() const;
			public: Token::SubType GetSubType();
			public: unsigned char GetParameterCount();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual void InsertColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void DeleteColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void InsertRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual void DeleteRow(unsigned short nWorksheet, unsigned short nRow);
			public: static const char* GetTypeName(Type eType);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SumToken : public Token
		{
			public: SumToken(unsigned char nParameterCount);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StringToken : public Token
		{
			protected: InternalString* m_sString = 0;
			public: StringToken(const char* sxString);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual ~StringToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SpaceToken : public Token
		{
			public: enum SpaceType
			{
				TYPE_SPACE_BEFORE_BASE_EXPRESSION = 0x00,
				TYPE_RETURN_BEFORE_BASE_EXPRESSION = 0x01,
				TYPE_SPACE_BEFORE_OPEN = 0x02,
				TYPE_RETURN_BEFORE_OPEN = 0x03,
				TYPE_SPACE_BEFORE_CLOSE = 0x04,
				TYPE_RETURN_BEFORE_CLOSE = 0x05,
				TYPE_SPACE_BEFORE_EXPRESSION = 0x06,
			};

			protected: SpaceType m_eSpaceType;
			protected: unsigned char m_nCount;
			public: SpaceToken(SpaceType eSpaceType, unsigned char nCount);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: SpaceToken::SpaceType GetSpaceType();
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OperatorToken : public Token
		{
			protected: InternalString* m_sOperator = 0;
			public: OperatorToken(const char* szOperator);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual ~OperatorToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class NumToken : public Token
		{
			protected: double m_fNumber;
			public: NumToken(double fNumber);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MissArgToken : public Token
		{
			public: MissArgToken();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IntToken : public Token
		{
			protected: unsigned short m_nInt;
			public: IntToken(unsigned short nInt);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ParseFunctionData
		{
			public: int m_nCount;
		};
		class ParseSpaceData
		{
			public: int m_nIndex;
			public: unsigned short m_cChar;
			public: int m_nCount;
			public: ParseSpaceData();
		};
		class Formula
		{
			protected: OwnedVector<Token*>* m_pTokenVector = 0;
			protected: Value* m_pValue = 0;
			protected: InternalString* m_sTemp = 0;
			public: Formula(const char* szFormula, WorksheetImplementation* pWorksheetImplementation);
			public: Formula(Vector<ParsedExpressionRecord*>* pParsedExpressionRecordVector, WorkbookGlobals* pWorkbookGlobals);
			public: Formula(OwnedVector<ParsedExpressionRecord*>* pParsedExpressionRecordVector, WorkbookGlobals* pWorkbookGlobals);
			public: unsigned short GetNumToken();
			public: const Token* GetTokenByIndex(unsigned short nIndex);
			public: const Value* Evaluate(WorksheetImplementation* pWorksheetImplementation, unsigned short nDepth);
			public: const char* ToString(WorksheetImplementation* pWorksheetImplementation);
			public: void ToRgce(RgceStruct* pRgce, WorkbookGlobals* pWorkbookGlobals);
			public: void InsertColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: void DeleteColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: void InsertRow(unsigned short nWorksheet, unsigned short nRow);
			public: void DeleteRow(unsigned short nWorksheet, unsigned short nRow);
			public: bool ValidateForChart(WorksheetImplementation* pWorksheetImplementation);
			public: const char* ToChartString(WorksheetImplementation* pWorksheetImplementation);
			public: bool ValidateForChartName(WorksheetImplementation* pWorksheetImplementation);
			public: const char* ToChartNameString(WorksheetImplementation* pWorksheetImplementation);
			protected: unsigned char Parse(InternalString* sFormula, WorksheetImplementation* pWorksheetImplementation);
			protected: bool ParseFunction(InternalString* sFormula, const char* szFunction, ParseFunctionData* pData, WorksheetImplementation* pWorksheetImplementation);
			protected: bool ParseSpace(InternalString* sFormula, ParseSpaceData* pData);
			protected: bool ParseString(InternalString* sFormula);
			protected: bool ParseBool(InternalString* sFormula, WorksheetImplementation* pWorksheetImplementation);
			protected: bool ParseInt(InternalString* sFormula);
			protected: bool ParseFloat(InternalString* sFormula);
			protected: void InsertOperator(InternalString* sOperator, InternalString* sTrailingSpaces);
			public: virtual ~Formula();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CoordinateToken : public Token
		{
			protected: Coordinate* m_pCoordinate = 0;
			public: CoordinateToken(Coordinate* pCoordinate);
			public: Coordinate* GetCoordinate();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual ~CoordinateToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Coordinate3dToken : public Token
		{
			protected: Coordinate3d* m_pCoordinate3d = 0;
			public: Coordinate3dToken(Coordinate3d* pCoordinate3d);
			public: Coordinate3d* GetCoordinate3d();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual void InsertColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void DeleteColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void InsertRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual void DeleteRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual ~Coordinate3dToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoolToken : public Token
		{
			protected: bool m_bBool;
			public: BoolToken(bool bBool, bool bFunction);
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaToken : public Token
		{
			protected: Area* m_pArea = 0;
			public: AreaToken(Area* pArea);
			public: Area* GetArea();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual ~AreaToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Area3dToken : public Token
		{
			protected: Area3d* m_pArea3d = 0;
			public: Area3dToken(Area3d* pArea3d);
			public: Area3d* GetArea3d();
			public: virtual void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ParsedExpressionRecord* ToParsedExpression(WorkbookGlobals* pWorkbookGlobals);
			public: virtual bool Evaluate(WorksheetImplementation* pWorksheetImplementation, OwnedVector<Value*>* ppValueVector, unsigned short nDepth);
			public: virtual void InsertColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void DeleteColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: virtual void InsertRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual void DeleteRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual ~Area3dToken();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RedBlackNode
		{
			public: enum Color
			{
				COLOR_RED = 0,
				COLOR_BLACK,
			};

			public: virtual RedBlackNode::Color GetColor() const;
			public: virtual RedBlackNode* GetParent() const;
			public: virtual RedBlackNode* GetLeftChild() const;
			public: virtual RedBlackNode* GetRightChild() const;
			public: virtual RedBlackNode* GetChild(int nDirection) const;
			public: virtual void* GetObject() const;
		};
		class RedBlackNodeImplementation : public RedBlackNode
		{
			public: Color m_eColor;
			public: RedBlackNodeImplementation* m_pParent = 0;
			public: RedBlackNodeImplementation* m_pChild[2];
			public: void* m_pObject;
			public: RedBlackNodeImplementation(void* pObject);
			public: RedBlackNodeImplementation* GetUncle() const;
			public: RedBlackNodeImplementation* GetGrandparent() const;
			public: virtual RedBlackNode::Color GetColor() const;
			public: virtual RedBlackNode* GetParent() const;
			public: virtual RedBlackNode* GetLeftChild() const;
			public: virtual RedBlackNode* GetRightChild() const;
			public: virtual RedBlackNode* GetChild(int nDirection) const;
			public: virtual void* GetObject() const;
			public: virtual ~RedBlackNodeImplementation();
		};
		class RedBlackTree
		{
			public: typedef int (ComparisonCallback)(const void* pObjectA, const void* pObjectB);
			protected: RedBlackTreeImplementation* m_pImpl = 0;
			public: RedBlackTree(ComparisonCallback* pComparisonCallback);
			public: virtual ~RedBlackTree();
			public: bool AddObject(void* pObject);
			public: bool DeleteObject(void* pObject);
			public: RedBlackNode* GetRootNode();
			public: RedBlackNode* GetNode(void* pObject);
		};
		class RedBlackTreeImplementation
		{
			public: RedBlackNodeImplementation* m_pRootNode = 0;
			public: RedBlackTree::ComparisonCallback* m_pComparisonCallback;
			public: void RecursiveDelete(RedBlackNodeImplementation* pNode);
			public: void Rotate(RedBlackNodeImplementation* pNode, int nDirection);
			public: virtual ~RedBlackTreeImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		template <class T>
		class TableElement
		{
			public: int m_nColumn;
			public: int m_nRow;
			public: T m_xObject = 0;
			public: virtual ~TableElement()
			{
				if (m_xObject != 0)
					{
						delete m_xObject;
						m_xObject = 0;
					}
			}

		};
		template <class T>
		class Table
		{
			protected: OwnedVector<TableElement<T>*>* m_pElementVector = 0;
			public: Table()
			{
				m_pElementVector = new OwnedVector<TableElement<T>*>();
			}

			public: void Set(int nColumn, int nRow, T xObject)
			{
				TableElement<T>* pElement = GetOrCreate(nColumn, nRow);
				pElement->m_xObject = xObject;
			}

			public: int GetIndex(int nColumn, int nRow)
			{
				int i = 0;
				while (true)
				{
					if (i >= m_pElementVector->GetSize())
						return i;
					TableElement<T>* pElement = m_pElementVector->Get(i);
					if (pElement->m_nRow > nRow)
						return i;
					if (pElement->m_nRow == nRow)
						if (pElement->m_nColumn >= nColumn)
							return i;
					i++;
				}
			}

			public: TableElement<T>* Get(int nColumn, int nRow)
			{
				int nIndex = GetIndex(nColumn, nRow);
				TableElement<T>* pElement = GetByIndex(nIndex);
				if (pElement == 0 || pElement->m_nColumn != nColumn || pElement->m_nRow != nRow)
					return 0;
				return pElement;
			}

			public: TableElement<T>* GetByIndex(int nIndex)
			{
				if (nIndex >= m_pElementVector->GetSize())
					return 0;
				return m_pElementVector->Get(nIndex);
			}

			public: TableElement<T>* GetOrCreate(int nColumn, int nRow)
			{
				int nIndex = GetIndex(nColumn, nRow);
				TableElement<T>* pElement = GetByIndex(nIndex);
				if (pElement == 0 || pElement->m_nColumn != nColumn || pElement->m_nRow != nRow)
				{
					TableElement<T>* pOwnedElement = new TableElement<T>();
					pOwnedElement->m_nColumn = nColumn;
					pOwnedElement->m_nRow = nRow;
					pOwnedElement->m_xObject = 0;
					pElement = pOwnedElement;
					{
						NumberDuck::Secret::TableElement<T>* __828927520 = pOwnedElement;
						pOwnedElement = 0;
						m_pElementVector->Insert(nIndex, __828927520);
					}
					if (pOwnedElement) delete pOwnedElement;
				}
				return pElement;
			}

			public: void Erase(int nIndex)
			{
				m_pElementVector->Erase(nIndex);
			}

			public: T PopBack()
			{
				TableElement<T>* pElement = m_pElementVector->PopBack();
				T xObject;
				{
					T __3920382863 = pElement->m_xObject;
					pElement->m_xObject = 0;
					xObject = __3920382863;
				}
				{
					delete pElement;
					pElement = 0;
				}
				{
					T __1841677085 = xObject;
					xObject = 0;
					{
						if (pElement) delete pElement;
						return __1841677085;
					}
				}
			}

			public: int GetSize()
			{
				return m_pElementVector->GetSize();
			}

			public: virtual ~Table()
			{
				if (m_pElementVector) delete m_pElementVector;
			}

		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PngImageInfo
		{
			public: int m_nWidth;
			public: int m_nHeight;
		};
		class PngLoader
		{
			protected: PngImageInfo* m_pImageInfo = 0;
			public: PngImageInfo* Load(Blob* pBlob);
			public: virtual ~PngLoader();
		};
	}
}

namespace NumberDuck
{
	class MD4
	{
		protected: static const int BLOCK_SIZE = 64;
		protected: unsigned int m_nBuffer[4];
		public: void Process(BlobView* pBlobView);
		public: void BlobWrite(BlobView* pBlobView);
		protected: void Reset();
		protected: void ProcessChunk(BlobView* pBlobView);
		protected: unsigned int FF(unsigned int a, unsigned int b, unsigned int c, unsigned int d, unsigned int x, int s);
		protected: unsigned int GG(unsigned int a, unsigned int b, unsigned int c, unsigned int d, unsigned int x, int s);
		protected: unsigned int HH(unsigned int a, unsigned int b, unsigned int c, unsigned int d, unsigned int x, int s);
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class JpegImageInfo
		{
			public: int m_nWidth;
			public: int m_nHeight;
		};
		class JpegLoader
		{
			protected: JpegImageInfo* m_pImageInfo = 0;
			public: JpegImageInfo* Load(Blob* pBlob);
			public: virtual ~JpegLoader();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XlsxWorksheet : public Worksheet
		{
			public: XlsxWorksheet(Workbook* pWorkbook);
			public: bool Parse(XlsxWorkbookGlobals* pWorkbookGlobals, XmlNode* pWorksheetNode);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XlsxWorkbookGlobals : public WorkbookGlobals
		{
			public: static bool LoadVector(Vector<XmlNode*>* pVector, XmlNode* pStyleSheetNode, const char* szParent, const char* szChild);
			public: static XmlNode* GetElement(Vector<XmlNode*>* pVector, int nIndex);
			public: static unsigned int ParseColor(XmlNode* pColorNode);
			public: static bool ApplyFont(XmlNode* pFontNode, Style* pStyle);
			public: static bool ApplyBorderLine(XmlNode* pSideElement, Line* pLine);
			public: static bool ApplyBorder(XmlNode* pBorderElement, Style* pStyle);
			public: static bool ApplyFill(XmlNode* pFillElement, Style* pStyle);
			public: static bool ParseStyles(WorkbookGlobals* pWorkbookGlobals, XmlNode* pStyleSheetNode);
			public: static bool SubParseStyles(WorkbookGlobals* pWorkbookGlobals, XmlNode* pStyleSheetNode, Vector<XmlNode*>* pNumFmtVector, Vector<XmlNode*>* pFontVector, Vector<XmlNode*>* pFillVector, Vector<XmlNode*>* pBorderVector, Vector<XmlNode*>* pCellStyleXfVector, Vector<XmlNode*>* pCellXfVector, Vector<XmlNode*>* pCellStyleVector);
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColumnInfo
		{
			public: unsigned short m_nWidth;
			public: bool m_bHidden;
		};
		class RowInfo
		{
			public: unsigned short m_nHeight;
		};
		class Coordinate
		{
			public: Coordinate();
			public: Coordinate(unsigned short nX, unsigned short nY, bool bXRelative = true, bool bYRelative = true);
			public: Coordinate* CreateClone();
			public: unsigned short m_nX;
			public: unsigned short m_nY;
			public: bool m_bXRelative;
			public: bool m_bYRelative;
		};
		class WorksheetRange
		{
			public: WorksheetRange(unsigned short nFirst, unsigned short nLast);
			public: WorksheetRange* CreateClone();
			public: unsigned short m_nFirst;
			public: unsigned short m_nLast;
		};
		class Coordinate3d
		{
			public: unsigned short m_nWorksheetFirst;
			public: unsigned short m_nWorksheetLast;
			public: Coordinate* m_pCoordinate = 0;
			public: Coordinate3d(unsigned short nWorksheetFirst, unsigned short nWorksheetLast, Coordinate* pCoordinate);
			public: Coordinate3d* CreateClone();
			public: void ToString(WorksheetImplementation* pWorksheetImplementation, InternalString* sOut);
			public: virtual ~Coordinate3d();
		};
		class Area
		{
			public: Coordinate* m_pTopLeft = 0;
			public: Coordinate* m_pBottomRight = 0;
			public: Area(Coordinate* pTopLeft, Coordinate* pBottomRight);
			public: Area* CreateClone();
			public: virtual ~Area();
		};
		class Area3d
		{
			public: WorksheetRange* m_pWorksheetRange = 0;
			public: Area* m_pArea = 0;
			public: Area3d(unsigned short nWorksheetFirst, unsigned short nWorksheetLast, Area* pArea);
			public: Area3d* CreateClone();
			public: virtual ~Area3d();
		};
		class WorksheetImplementation
		{
			public: Workbook* m_pWorkbook = 0;
			public: InternalString* m_sName = 0;
			public: Worksheet::Orientation m_eOrientation;
			public: bool m_bShowGridlines;
			public: bool m_bPrintGridlines;
			public: Table<Cell*>* m_pCellTable = 0;
			public: Table<ColumnInfo*>* m_pColumnInfoTable = 0;
			public: Table<RowInfo*>* m_pRowInfoTable = 0;
			public: unsigned short m_nDefaultRowHeight;
			public: OwnedVector<Picture*>* m_pPictureVector = 0;
			public: OwnedVector<Chart*>* m_pChartVector = 0;
			public: OwnedVector<MergedCell*>* m_pMergedCellVector = 0;
			public: Worksheet* m_pWorksheet = 0;
			public: WorksheetImplementation(Workbook* pWorkbook, Worksheet* pWorksheet);
			public: virtual ~WorksheetImplementation();
			public: static unsigned short TwipsToPixels(unsigned short nTwips);
			public: static unsigned short PixelsToTwips(unsigned short nPixels);
			public: static void CoordinateToAddress(Coordinate* pCoordinate, InternalString* sOut);
			public: static void AreaToAddress(Area* pArea, InternalString* sOut);
			public: static Coordinate* AddressToCoordinate(const char* szAddress);
			public: static Area* AddressToArea(const char* szAddress);
			public: void WorksheetRangeToAddress(unsigned short nFirst, unsigned short nLast, InternalString* sOut);
			public: void Area3dToAddress(Area3d* pArea3d, InternalString* sOut);
			public: WorksheetRange* ParseWorksheetRange(InternalString* sString);
			public: Coordinate3d* ParseCoordinate3d(InternalString* sString);
			public: Area3d* ParseArea3d(InternalString* sString);
			public: ColumnInfo* GetColumnInfo(unsigned short nColumn);
			public: ColumnInfo* GetOrCreateColumnInfo(unsigned short nColumn);
			public: RowInfo* GetRowInfo(unsigned short nRow);
			public: RowInfo* GetOrCreateRowInfo(unsigned short nRow);
			public: Workbook* GetWorkbook();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WorkbookImplementation
		{
			public: WorkbookGlobals* m_pWorkbookGlobals = 0;
			public: OwnedVector<Worksheet*>* m_pWorksheetVector = 0;
			public: virtual ~WorkbookImplementation();
		};
	}
}

namespace NumberDuck
{
	class Workbook
	{
		public: enum License
		{
			AGPL,
			Commercial,
		};

		public: enum FileType
		{
			XLS,
			XLSX,
		};

		public: Secret::WorkbookImplementation* m_pImpl = 0;
		public: Workbook(License eLicense);
		public: void Clear();
		public: bool Load(const char* szFileName);
		public: bool Save(const char* szFileName, FileType eFileType);
		public: unsigned int GetNumWorksheet();
		public: Worksheet* GetWorksheetByIndex(unsigned int nIndex);
		public: Worksheet* CreateWorksheet();
		public: void PurgeWorksheet(unsigned int nIndex);
		public: unsigned int GetNumStyle();
		public: Style* GetStyleByIndex(unsigned int nIndex);
		public: Style* GetDefaultStyle();
		public: Style* CreateStyle();
		public: virtual ~Workbook();
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class ValueImplementation
		{
			public: Value::Type m_eType;
			public: InternalString* m_sString = 0;
			public: double m_fFloat;
			public: bool m_bBoolean;
			public: Formula* m_pFormula = 0;
			public: Worksheet* m_pWorksheet = 0;
			public: Area* m_pArea = 0;
			public: Area3d* m_pArea3d = 0;
			public: Value* m_pValue = 0;
			public: ValueImplementation();
			public: static Value* CreateFloatValue(double fFloat);
			public: static Value* CreateStringValue(const char* szString);
			public: static Value* CreateBooleanValue(const bool bBoolean);
			public: static Value* CreateErrorValue();
			public: static Value* CreateAreaValue(Area* pArea);
			public: static Value* CreateArea3dValue(Area3d* pArea3d);
			public: static Value* CopyValue(const Value* pValue);
			public: void Clear();
			public: void SetString(const char* szString);
			public: void SetFloat(double fFloat);
			public: void SetBoolean(bool bBoolean);
			public: void SetFormulaFromString(const char* szFormula, Worksheet* pWorksheet);
			public: void SetFormula(Formula* pFormula, Worksheet* pWorksheet);
			public: virtual ~ValueImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StyleImplementation
		{
			public: Font* m_pFont = 0;
			public: Style::HorizontalAlign m_eHorizontalAlign;
			public: Style::VerticalAlign m_eVerticalAlign;
			public: Color* m_pBackgroundColor = 0;
			public: Style::FillPattern m_eFillPattern;
			public: Color* m_pFillPatternColor = 0;
			public: Line* m_pTopBorderLine = 0;
			public: Line* m_pRightBorderLine = 0;
			public: Line* m_pBottomBorderLine = 0;
			public: Line* m_pLeftBorderLine = 0;
			public: InternalString* m_sFormat = 0;
			public: unsigned int m_nFormatIndex;
			public: StyleImplementation();
			public: virtual ~StyleImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SharedString
		{
			public: InternalString* m_sString = 0;
			public: int m_nChecksum;
			public: int m_nIndex;
			public: virtual ~SharedString();
		};
		class SharedStringContainer
		{
			protected: OwnedVector<SharedString*>* m_pSharedStringVector = 0;
			protected: Vector<SharedString*>* m_pSharedStringSortedVector = 0;
			public: SharedStringContainer();
			public: virtual ~SharedStringContainer();
			public: void Clear();
			public: const char* Get(int nIndex);
			public: int GetIndex(const char* sxString);
			public: int Push(const char* sxString);
			public: int GetSize();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesImplementation
		{
			public: Worksheet* m_pWorksheet = 0;
			public: Formula* m_pNameFormula = 0;
			public: Formula* m_pValuesFormula = 0;
			public: Line* m_pLine = 0;
			public: Fill* m_pFill = 0;
			public: Marker* m_pMarker = 0;
			public: SeriesImplementation(Worksheet* pWorksheet, Formula* pValuesFormula);
			public: void SetNameFormula(Formula* pFormula);
			public: void SetValuesFormula(Formula* pFormula);
			public: void SetClassicStyle(Chart::Type eChartType, unsigned short nIndex);
			public: virtual ~SeriesImplementation();
		};
	}
}

namespace NumberDuck
{
	class Series
	{
		public: Secret::SeriesImplementation* m_pImpl = 0;
		public: Series(Worksheet* pWorksheet, Secret::Formula* pValuesFormula);
		public: const char* GetName();
		public: bool SetName(const char* szName);
		public: const char* GetValues();
		public: bool SetValues(const char* szValues);
		public: Line* GetLine();
		public: Fill* GetFill();
		public: Marker* GetMarker();
		public: virtual ~Series();
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class PictureImplementation
		{
			public: unsigned int m_nX;
			public: unsigned int m_nY;
			public: unsigned int m_nSubX;
			public: unsigned int m_nSubY;
			public: unsigned int m_nWidth;
			public: unsigned int m_nHeight;
			public: InternalString* m_sUrl = 0;
			public: Blob* m_pBlob = 0;
			public: Picture::Format m_eFormat;
			public: PictureImplementation(Blob* pBlob, Picture::Format eFormat);
			public: virtual ~PictureImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MergedCellImplementation
		{
			public: unsigned int m_nX;
			public: unsigned int m_nY;
			public: unsigned int m_nWidth;
			public: unsigned int m_nHeight;
		};
	}
}

namespace NumberDuck
{
	class MergedCell
	{
		protected: Secret::MergedCellImplementation* m_pImpl = 0;
		public: MergedCell(unsigned int nX, unsigned int nY, unsigned int nWidth, unsigned int nHeight);
		public: unsigned int GetX() const;
		public: void SetX(unsigned int nX);
		public: unsigned int GetY() const;
		public: void SetY(unsigned int nY);
		public: unsigned int GetWidth() const;
		public: void SetWidth(unsigned int nWidth);
		public: unsigned int GetHeight() const;
		public: void SetHeight(unsigned int nHeight);
		public: virtual ~MergedCell();
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class MarkerImplementation
		{
			public: Marker::Type m_eType;
			public: Color* m_pFillColor = 0;
			public: Color* m_pBorderColor = 0;
			public: int m_nSize;
			public: MarkerImplementation();
			public: bool Equals(const MarkerImplementation* pMarkerImplementation) const;
			public: Marker::Type GetType() const;
			public: void SetType(Marker::Type eType);
			public: Color* GetFillColor(bool bCreateIfMissing);
			public: void SetFillColor(const Color* pColor);
			public: void ClearFillColor();
			public: Color* GetBorderColor(bool bCreateIfMissing);
			public: void SetBorderColor(const Color* pColor);
			public: void ClearBorderColor();
			public: int GetSize() const;
			public: void SetSize(int nSize);
			public: virtual ~MarkerImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineImplementation
		{
			public: Line::Type m_eType;
			public: Color* m_pColor = 0;
			public: LineImplementation();
			public: bool Equals(const LineImplementation* pLineImplementation) const;
			public: Line::Type GetType() const;
			public: void SetType(Line::Type eType);
			public: virtual ~LineImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LegendImplementation
		{
			public: bool m_bHidden;
			public: Line* m_pBorderLine = 0;
			public: Fill* m_pFill = 0;
			public: LegendImplementation();
			public: virtual ~LegendImplementation();
		};
	}
}

namespace NumberDuck
{
	class Legend
	{
		protected: Secret::LegendImplementation* m_pImpl = 0;
		public: Legend();
		public: bool GetHidden() const;
		public: void SetHidden(bool bHidden);
		public: Line* GetBorderLine();
		public: Fill* GetFill();
		public: virtual ~Legend();
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontImplementation
		{
			public: InternalString* m_sName = 0;
			public: int m_nSizeTwips;
			public: Color* m_pColor = 0;
			public: bool m_bBold;
			public: bool m_bItalic;
			public: Font::Underline m_eUnderline;
			public: FontImplementation();
			public: virtual ~FontImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FillImplementation
		{
			public: Fill::Type m_eType;
			public: Color* m_pForegroundColor = 0;
			public: Color* m_pBackgroundColor = 0;
			public: FillImplementation();
			public: bool Equals(const FillImplementation* pFillImplementation) const;
			public: Fill::Type GetType() const;
			public: void SetType(Fill::Type eType);
			public: virtual ~FillImplementation();
		};
	}
}

namespace NumberDuck
{
	class Color
	{
		protected: unsigned char m_nRed;
		protected: unsigned char m_nGreen;
		protected: unsigned char m_nBlue;
		public: Color(unsigned char nRed, unsigned char nGreen, unsigned char nBlue);
		public: bool Equals(const Color* pColor) const;
		public: unsigned char GetRed() const;
		public: unsigned char GetGreen() const;
		public: unsigned char GetBlue() const;
		public: void Set(unsigned char nRed, unsigned char nGreen, unsigned char nBlue);
		public: void SetRed(unsigned char nRed);
		public: void SetGreen(unsigned char nGreen);
		public: void SetBlue(unsigned char nBlue);
		public: void SetFromColor(const Color* pColor);
		public: void SetFromRgba(unsigned int nRgba);
		public: unsigned int GetRgba() const;
	};
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartImplementation
		{
			public: Worksheet* m_pWorksheet = 0;
			public: unsigned int m_nX;
			public: unsigned int m_nY;
			public: unsigned int m_nSubX;
			public: unsigned int m_nSubY;
			public: unsigned int m_nWidth;
			public: unsigned int m_nHeight;
			public: Chart::Type m_eType;
			public: Formula* m_pCategoriesFormula = 0;
			public: InternalString* m_sTitle = 0;
			public: InternalString* m_sHorizontalAxisLabel = 0;
			public: InternalString* m_sVerticalAxisLabel = 0;
			public: Legend* m_pLegend = 0;
			public: Line* m_pFrameBorderLine = 0;
			public: Fill* m_pFrameFill = 0;
			public: Line* m_pPlotBorderLine = 0;
			public: Fill* m_pPlotFill = 0;
			public: Line* m_pHorizontalAxisLine = 0;
			public: Line* m_pHorizontalGridLine = 0;
			public: Line* m_pVerticalAxisLine = 0;
			public: Line* m_pVerticalGridLine = 0;
			public: OwnedVector<Series*>* m_pSeriesVector = 0;
			public: ChartImplementation(Worksheet* pWorksheet, Chart::Type eType);
			public: Series* CreateSeries(Formula* pValuesFormula);
			public: void PurgeSeries(int nIndex);
			public: void SetCategoriesFormula(Formula* pCategoriesFormula);
			public: void SetClassicStyle();
			public: void InsertColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: void DeleteColumn(unsigned short nWorksheet, unsigned short nColumn);
			public: void InsertRow(unsigned short nWorksheet, unsigned short nRow);
			public: void DeleteRow(unsigned short nWorksheet, unsigned short nRow);
			public: virtual ~ChartImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellImplementation
		{
			public: Worksheet* m_pWorksheet = 0;
			public: Value* m_pValue = 0;
			public: Style* m_pStyle = 0;
			public: Workbook* GetWorkbook();
			public: void SetFormula(Formula* pFormula);
			public: virtual ~CellImplementation();
		};
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		template <class T>
		class OwnedVector
		{
			protected: Vector<T>* m_pVector = 0;
			public: OwnedVector()
			{
				m_pVector = new Vector<T>();
			}

			public: virtual ~OwnedVector()
			{
				Clear();
				if (m_pVector) delete m_pVector;
			}

			public: T PushBack(T xObject)
			{
				m_pVector->PushBack(xObject);
				return m_pVector->Get(m_pVector->GetSize() - 1);
			}

			public: int GetSize()
			{
				return m_pVector->GetSize();
			}

			public: T Get(int nIndex)
			{
				return m_pVector->Get(nIndex);
			}

			public: void Clear()
			{
				while (m_pVector->GetSize() > 0)
				{
					T pTemp = m_pVector->PopBack();
					{
						delete pTemp;
						pTemp = 0;
					}
				}
			}

			public: void Insert(int nIndex, T pObject)
			{
				m_pVector->Insert(nIndex, pObject);
			}

			public: T Remove(int nIndex)
			{
				T pTemp = m_pVector->Get(nIndex);
				m_pVector->Erase(nIndex);
				return pTemp;
			}

			public: void Erase(int nIndex)
			{
				{
					delete m_pVector->Get(nIndex);
				}
				m_pVector->Erase(nIndex);
			}

			public: T PopBack()
			{
				return m_pVector->PopBack();
			}

			public: T PopFront()
			{
				return m_pVector->PopFront();
			}

		};
	}
}

#pragma once

namespace NumberDuck
{
	class BlobView;

	class Blob
	{
		public:
			Blob(bool bAutoResize);
			~Blob();

			bool Load(const char* szFileName);
			bool Save(const char* szFileName);

			void Resize(int nSize, bool bAutoResize);
			int GetSize();

			unsigned int GetMsoCrc32();

			BlobView* GetBlobView();

			bool Equal(Blob* pOther);

			// to hide
			void PackData(unsigned char* pData, int nOffset, int nSize);
			void UnpackData(unsigned char* pData, int nOffset, int nSize);
			unsigned char* GetData();

		protected:
			friend class BlobView;

			static const int DEFAULT_SIZE = 1024 * 32;

			int m_nBufferSize;
			unsigned char* m_pBuffer;
			int m_nSize;
			bool m_bAutoResize;
			BlobView* m_pBlobView;
	};

	class BlobView
	{
		public:
			BlobView(Blob* pBlob, int nStart, int nEnd);
			~BlobView();

			unsigned char GetChecksum();

			int GetSize();

			void PackUint8(unsigned char n);
			void PackUint16(unsigned short n);
			void PackUint32(unsigned int n);

			void PackInt8(signed char n);
			void PackInt16(signed short n);
			void PackInt32(signed int n);

			void PackDouble(double n);

			unsigned char UnpackUint8();
			unsigned short UnpackUint16();
			unsigned int UnpackUint32();

			signed char UnpackInt8();
			signed short UnpackInt16();
			signed int UnpackInt32();

			double UnpackDouble();

			signed int UnpackInt32At(int nOffset);

			int GetStart();
			int GetEnd();

			void SetEnd(int nEnd);

			int GetOffset();
			void SetOffset(int nOffset);

			void Pack(BlobView* pBlobView, int nSize);
			void PackAt(int nOffset, BlobView* pBlobView, int nSize);

			void Unpack(BlobView* pBlobView, int nSize);
			void UnpackAt(int nOffset, BlobView* pBlobView, int nSize);

			Blob* GetBlob();

			//protected:
			// hax for now

			void PackData(unsigned char* pData, int nSize);
			void UnpackData(unsigned char* pData, int nSize);

			void PackDataAt(int nOffset, unsigned char* pData, int nSize);
			void UnpackDataAt(int nOffset, unsigned char* pData, int nSize);

		protected:
			int m_nStart;
			int m_nEnd;
			int m_nOffset;

			Blob* m_pBlob;
	};
}


