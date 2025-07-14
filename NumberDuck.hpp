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
		class WorkbookImplementation;
		template <class T>
		class OwnedVector;
		class WorkbookGlobals;
		class CompoundFile;
		class Stream;
		class BiffRecord;
		class BofRecord;
		class BiffWorkbookGlobals;
		class BiffWorksheet;
		class BiffRecordContainer;
		class SeriesImplementation;
		class MergedCellImplementation;
		class LegendImplementation;
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

		public: Secret::ValueImplementation* m_pImpl;
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
	class Cell
	{
		public: Secret::CellImplementation* m_pImpl;
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
	class Style
	{
		public: Secret::StyleImplementation* m_pImplementation;
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
	class Line
	{
		protected: Secret::LineImplementation* m_pImpl;
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
	class Marker
	{
		public: Secret::MarkerImplementation* m_pImpl;
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
	class Fill
	{
		protected: Secret::FillImplementation* m_pImpl;
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
	class Font
	{
		public: Secret::FontImplementation* m_pImpl;
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

		public: Secret::ChartImplementation* m_pImpl;
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
	class Picture
	{
		public: Secret::PictureImplementation* m_pImplementation;
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
	class Worksheet
	{
		public: Secret::WorksheetImplementation* m_pImpl;
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
		public: MD4();
	};
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

		public: Secret::WorkbookImplementation* m_pImpl;
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
	class Series
	{
		public: Secret::SeriesImplementation* m_pImpl;
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
	class MergedCell
	{
		protected: Secret::MergedCellImplementation* m_pImpl;
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
	class Legend
	{
		protected: Secret::LegendImplementation* m_pImpl;
		public: Legend();
		public: bool GetHidden() const;
		public: void SetHidden(bool bHidden);
		public: Line* GetBorderLine();
		public: Fill* GetFill();
		public: virtual ~Legend();
	};
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


