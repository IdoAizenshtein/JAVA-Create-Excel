package bizibox.exportsExcelMail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.*;
import java.util.Calendar;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import java.io.*;

import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import javax.ws.rs.*;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

@Path("/excelExports")
public class ExcelExports extends SendFileEmail {
    Response.ResponseBuilder rb = null;
    private static final String XML_ENCODING = "UTF-8";
    private static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    private static final int UNIT_OFFSET_LENGTH = 7;

    private enum TypesAlign {
        center,
        right,
        left
    }

    static int getWidth(XSSFSheet sheet, int col) {
        int width = sheet.getColumnWidth(col);
        if (width == sheet.getDefaultColumnWidth()) {
            width = (short) (width * 256);
        }
        return width;
    }

    private enum TypesVertical {
        bottom,
        top,
        justify,
        center
    }

    @POST
    @Path("/download")
    @Consumes(MediaType.APPLICATION_FORM_URLENCODED)
    @Produces("application/vnd.ms-excel")
    public Response downloadXLS(@FormParam("data") String incomingData) throws IOException {
        return process(incomingData);
    }

    @POST
    @Path("/sendMail")
    @Consumes({MediaType.APPLICATION_JSON})
    @Produces({MediaType.TEXT_PLAIN})
    public Response sendMailer(String inputJsonObj) throws IOException {
        return process(inputJsonObj);
    }

    public Response process(String incomingData) throws IOException {
        try {
            JSONParser parser = new JSONParser();
            Object objExcel = parser.parse(incomingData);
            JSONObject objValMain = (JSONObject) objExcel;
            return processFile(objValMain);
        } catch (ParseException pe) {
            // System.out.println("position: " + pe.getPosition());
            return Response.status(200).entity("[]").build();
        }
    }

    public Response processFile(JSONObject objValMain) throws IOException {
        String pathLocal = System.getProperty("catalina.base");
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("bizibox");

        JSONArray rows = (JSONArray) objValMain.get("rows");
        JSONArray stylesAll = (JSONArray) objValMain.get("styles");
        JSONArray mergeCells = (JSONArray) objValMain.get("mergeCells");
        JSONArray widthCells = (JSONArray) objValMain.get("widthCells");
        JSONArray senderMail = (JSONArray) objValMain.get("senderMail");
        JSONObject heightRows = (JSONObject) objValMain.get("heightRows");

        // System.out.println("1: ");

        String xmlMerge = "";
        int sumOfMerge = mergeCells.size();
        if (sumOfMerge > 0) {
            Iterator itr = mergeCells.iterator();
            xmlMerge += "<mergeCells count=\"" + sumOfMerge + "\">";
            while (itr.hasNext()) {
                Object element = itr.next();
                String mergeCellStr = element.toString();
                xmlMerge += "<mergeCell ref=\"" + mergeCellStr + "\"/>";
            }
            xmlMerge += "</mergeCells>";
        }

        String xmlWidth = "";
        int sumOfwidth = widthCells.size();
        if (sumOfwidth > 0) {
            xmlWidth += "<cols>";
            Iterator itrators = widthCells.iterator();
            int idxWidth = 0;
            while (itrators.hasNext()) {
                idxWidth += 1;
                Object elements = itrators.next();
                String widthCellStr = elements.toString();
                xmlWidth += "<col min=\"" + idxWidth + "\" bestFit=\"1\" max=\"" + idxWidth + "\" width=\"" + widthCellStr + "\" customWidth=\"1\"/>";
            }
            xmlWidth += "</cols>";
        }

        Map<String, XSSFCellStyle> styles = createStyles(wb, stylesAll);
        String sheetRef = sheet.getPackagePart().getPartName().getName();

        File file = new File(pathLocal + "/template.xlsx");
        file.setWritable(true, false);

        FileOutputStream os = new FileOutputStream(file);
        wb.write(os);
        os.close();

        File tmp = File.createTempFile("sheet", ".xml");
        Writer fw = new OutputStreamWriter(new FileOutputStream(tmp), XML_ENCODING);
        try {
            generate(fw, styles, rows, xmlMerge, xmlWidth, heightRows);
            fw.close();
            File fileXlsx = new File(pathLocal + "/excel.xlsx");
            fileXlsx.setWritable(true, false);

            FileOutputStream out = new FileOutputStream(fileXlsx);
            substitute(file, tmp, sheetRef.substring(1), out);
            out.close();
            wb.close();

            int mailerExist = senderMail.size();
            if (mailerExist > 0) {
                // System.out.println("2: ");
                JSONObject objValMail = (JSONObject) senderMail.get(0);
                Object title = objValMail.get("title");
                Object name_company = objValMail.get("name_company");
                Object name_roh = objValMail.get("name_roh");
                Object name_doch = objValMail.get("name_doch");
                Object toAddressMail = objValMail.get("toAddressMail");

                String titleStr = title.toString();
                String name_companyStr = name_company.toString();
                String name_rohStr = name_roh.toString();
                String name_dochStr = name_doch.toString();
                String toAddressMailStr = toAddressMail.toString();

                String isSend = sender(fileXlsx, titleStr, name_companyStr, name_rohStr, name_dochStr, toAddressMailStr);
                String isSender = String.valueOf(isSend);
                return Response.status(200).entity(isSender).build();
            } else {
                //System.out.println("3: ");
                rb = Response.ok(fileXlsx);
                rb.header("Content-Disposition", "attachment; filename=excel.xlsx");
                return rb.status(200).build();
            }
        } catch (Exception e) {
            return Response.status(200).entity("[]").build();
        }
    }

    private static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook wb, JSONArray stylesAll) {
        Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
        int numerOfArray = stylesAll.size();
        for (int i = 0; i < numerOfArray; i++) {
            JSONObject objVal = (JSONObject) stylesAll.get(i);
            Object fontSize = objVal.get("fontSize");
            Object bold = objVal.get("bold");
            Object fontItalic = objVal.get("fontItalic");
            Object fontUnderline = objVal.get("fontUnderline");
            Object color = objVal.get("color");
            Object fillForegroundColor = objVal.get("fillForegroundColor");
            Object alignment = objVal.get("alignment");
            Object verticalAlignment = objVal.get("verticalAlignment");
            Object borderRight = objVal.get("borderRight");
            Object borderRightColor = objVal.get("borderRightColor");
            Object borderLeft = objVal.get("borderLeft");
            Object borderLeftColor = objVal.get("borderLeftColor");
            Object borderTop = objVal.get("borderTop");
            Object borderTopColor = objVal.get("borderTopColor");
            Object borderBottom = objVal.get("borderBottom");
            Object borderBottomColor = objVal.get("borderBottomColor");
            Object typesCell = objVal.get("type");
            Object typesCellNumber = objVal.get("typesCellNumber");

            String typesCellName = typesCell.toString();
            XSSFCellStyle style = wb.createCellStyle();

            if (typesCellNumber == (Boolean) true) {
                XSSFDataFormat fmt = wb.createDataFormat();
                style.setDataFormat(fmt.getFormat("_ [$₪-40D] * #,##0.00_ ;_ [$₪-40D] * -#,##0.00_ ;_ [$₪-40D] * \"-\"??_ ;_ @_ "));
            }

            //create font
            XSSFFont font = wb.createFont();
            boolean bolder = (Boolean) bold;
            font.setBold(bolder);

            boolean fontItalicBool = (Boolean) fontItalic;
            font.setItalic(fontItalicBool);

            if (fontUnderline != null) {
                boolean fontUnderlineBool = (Boolean) fontUnderline;
                if (fontUnderlineBool) {
                    font.setUnderline(XSSFFont.U_SINGLE);
                }
            }

            int fontSizePoint = Integer.parseInt(fontSize.toString());
            font.setFontHeight((double) fontSizePoint);
            font.setFontName("Arial");

            String rgbColor = color.toString();
            //convert hex to rgb color
            short[] colorsRgb = getRGB(rgbColor);
            font.setColor(new XSSFColor(new java.awt.Color(colorsRgb[0], colorsRgb[1], colorsRgb[2])));


            String bgColor = fillForegroundColor.toString();
            //convert hex to rgb color
            short[] bgRgb = getRGB(bgColor);
            style.setFillForegroundColor(new XSSFColor(new java.awt.Color(bgRgb[0], bgRgb[1], bgRgb[2])));
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);

            //set Align cell
            String alignmentString = alignment.toString();
            TypesAlign typesAlign = TypesAlign.valueOf(alignmentString);
            switch (typesAlign) {
                case center:
                    style.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case right:
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                case left:
                    style.setAlignment(HorizontalAlignment.LEFT);
                    break;
                default:
                    style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            }

            //set Vertical cell
            String verticalAlignmentString = verticalAlignment.toString();
            TypesVertical typesVertical = TypesVertical.valueOf(verticalAlignmentString);
            switch (typesVertical) {
                case bottom:
                    style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    break;
                case top:
                    style.setVerticalAlignment(VerticalAlignment.TOP);
                    break;
                case justify:
                    style.setVerticalAlignment(VerticalAlignment.JUSTIFY);
                    break;
                case center:
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    break;
                default:
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
            }

            //create borders
            //BORDER_DASH_DOT	Cell style with dash and dot
            //BORDER_DOTTED	Cell style with dotted border
            //BORDER_DASHED	Cell style with dashed border
            //BORDER_THICK	Cell style with thick border
            //BORDER_THIN	Cell style with thin border
            if (borderLeft == (Boolean) true) {
                //create border left
                style.setBorderLeft(XSSFCellStyle.BORDER_THIN);

                String rgbColor2 = borderLeftColor.toString();
                //convert hex to rgb color
                short[] colorsRgb2 = getRGB(rgbColor2);
                XSSFColor cellBorderColour2 = new XSSFColor(new java.awt.Color(colorsRgb2[0], colorsRgb2[1], colorsRgb2[2]));
                style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, cellBorderColour2);
            }
            if (borderTop == (Boolean) true) {
                //create border top
                style.setBorderTop(XSSFCellStyle.BORDER_THIN);

                String rgbColor3 = borderTopColor.toString();
                //convert hex to rgb color
                short[] colorsRgb3 = getRGB(rgbColor3);
                XSSFColor cellBorderColour3 = new XSSFColor(new java.awt.Color(colorsRgb3[0], colorsRgb3[1], colorsRgb3[2]));
                style.setBorderColor(XSSFCellBorder.BorderSide.TOP, cellBorderColour3);
            }
            if (borderRight == (Boolean) true) {
                //create border right
                style.setBorderRight(XSSFCellStyle.BORDER_THIN);

                String rgbColor1 = borderRightColor.toString();
                //convert hex to rgb color
                short[] colorsRgb1 = getRGB(rgbColor1);
                XSSFColor cellBorderColour1 = new XSSFColor(new java.awt.Color(colorsRgb1[0], colorsRgb1[1], colorsRgb1[2]));
                style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, cellBorderColour1);
            }
            if (borderBottom == (Boolean) true) {
                //create border bottom
                style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

                String rgbColor4 = borderBottomColor.toString();
                //convert hex to rgb color
                short[] colorsRgb4 = getRGB(rgbColor4);
                XSSFColor cellBorderColour4 = new XSSFColor(new java.awt.Color(colorsRgb4[0], colorsRgb4[1], colorsRgb4[2]));
                style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, cellBorderColour4);
            }

            //apply font
            style.setFont(font);
            styles.put(typesCellName, style);
        }
        return styles;
    }

    private static void generate(Writer out, Map<String, XSSFCellStyle> styles, JSONArray rows, String xmlMerge, String xmlWidth, JSONObject heightRows) throws Exception {
        SpreadsheetWriter sw = new SpreadsheetWriter(out);
        int[] arrWidth = null;
        if (xmlWidth.equals("")) {
            JSONObject row1 = (JSONObject) rows.get(0);
            JSONArray arrayAll1 = (JSONArray) row1.get("cell");
            int numerOfArrayRow1 = arrayAll1.size();
            arrWidth = new int[numerOfArrayRow1];
            Arrays.fill(arrWidth, 0);
        }
        try {
            int numerOfArray = rows.size();
            for (int i = 0; i < numerOfArray; i++) {
                String heightRo = "";
                if (heightRows != null) {
                    if (heightRows.get("row" + i) != null) {
                        heightRo = "customHeight=\"true\" ht=\"" + heightRows.get("row" + i) + "\"";
                    }
                }
                sw.insertRow(i, heightRo);

                JSONObject objValMain = (JSONObject) rows.get(i);
                JSONArray array1 = (JSONArray) objValMain.get("cell");
                int numerOfArray1 = array1.size();
                for (int i1 = 0; i1 < numerOfArray1; i1++) {
                    JSONObject objVal = (JSONObject) array1.get(i1);
                    Object textVal = objVal.get("val");
                    Object typeStyle = objVal.get("type");
                    Object typeTitle = objVal.get("title");

                    String val = textVal.toString().replaceAll("___", "&amp;");
                    String type = typeStyle.toString();
                    int styleIndex = styles.get(type).getIndex();
                    short isNumber = styles.get(type).getDataFormat();
                    //System.out.println(isNumber);
                    if (arrWidth != null && typeTitle == null) {
                        int lengStrVal = val.length();
                        if (isNumber == 165 && !val.equals("")) {
                            lengStrVal = lengStrVal + 7;
                        }
                        if (arrWidth[i1] < lengStrVal) {
                            arrWidth[i1] = lengStrVal;
                        }
                    }

                    if (isNumber == 165) {
                        if (val.equals("")) {
                            sw.createCell(i1, val, styleIndex);
                        } else {
                            try {
                                double valNum = Double.parseDouble(val);
                                sw.createCell(i1, valNum, styleIndex);
                            } catch (NumberFormatException nfe) {
                                sw.createCell(i1, val, styleIndex);
                            }
                        }
                    } else {
                        sw.createCell(i1, val, styleIndex);
                    }
                }
                sw.endRow();
            }
        } catch (Exception e) {
            if (arrWidth != null) {
                xmlWidth = getColWidth(xmlWidth, arrWidth);
            }
            sw.endSheet(xmlMerge);
            sw.beginSheet(xmlWidth);
        } finally {
            if (arrWidth != null) {
                xmlWidth = getColWidth(xmlWidth, arrWidth);
            }
            sw.endSheet(xmlMerge);
            sw.beginSheet(xmlWidth);
        }
    }

    private static int widthUnits2Pixel(short widthUnits) {
        int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR) * UNIT_OFFSET_LENGTH;
        int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
        pixels += Math.floor((float) offsetWidthUnits / ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));
        return pixels;
    }

    private static void substitute(File zipfile, File tmpfile, String entry, OutputStream out) throws IOException {
        ZipFile zip = ZipHelper.openZipFile(zipfile);
        try {
            ZipOutputStream zos = new ZipOutputStream(out);

            Enumeration<? extends ZipEntry> en = zip.entries();
            while (en.hasMoreElements()) {
                ZipEntry ze = en.nextElement();
                if (!ze.getName().equals(entry)) {
                    zos.putNextEntry(new ZipEntry(ze.getName()));
                    InputStream is = zip.getInputStream(ze);
                    copyStream(is, zos);
                    is.close();
                }
            }
            zos.putNextEntry(new ZipEntry(entry));
            InputStream is = new FileInputStream(tmpfile);
            copyStream(is, zos);
            is.close();

            zos.close();
        } finally {
            zip.close();
        }
    }

    private static void copyStream(InputStream in, OutputStream out) throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        while ((count = in.read(chunk)) >= 0) {
            out.write(chunk, 0, count);
        }
    }

    public static class SpreadsheetWriter {
        private final Writer _out;
        private int _rownum;
        private StringBuffer sBuffer;

        public SpreadsheetWriter(Writer out) {
            sBuffer = new StringBuffer();
            _out = out;
        }

        public void beginSheet(String xmlWidth) throws IOException {
            _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            _out.write("<sheetViews><sheetView colorId=\"64\" defaultGridColor=\"true\" rightToLeft=\"true\" showFormulas=\"false\" showGridLines=\"true\" showOutlineSymbols=\"true\" showRowColHeaders=\"true\" showZeros=\"true\" tabSelected=\"true\" topLeftCell=\"A1\" view=\"normal\" windowProtection=\"false\" workbookViewId=\"0\" zoomScale=\"100\" zoomScaleNormal=\"100\" zoomScalePageLayoutView=\"100\"> </sheetView></sheetViews>");
            _out.write("<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>");
            if (!xmlWidth.equals("")) {
                _out.write(xmlWidth);
            }
            _out.write("<sheetData>");
            _out.write(sBuffer.toString());
        }

        public void endSheet(String xmlMerge) throws IOException {
            sBuffer.append("</sheetData>");
            if (!xmlMerge.equals("")) {
                sBuffer.append(xmlMerge);
            }
            sBuffer.append("</worksheet>");
        }

        public void insertRow(int rownum, String heightRo) throws IOException {
            sBuffer.append("<row " + heightRo + " r=\"" + (rownum + 1) + "\">\n");
            this._rownum = rownum;
        }

        public void endRow() throws IOException {
            sBuffer.append("</row>\n");
        }

        public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
            String ref = new CellReference(_rownum, columnIndex).formatAsString();
            sBuffer.append("<c r=\"" + ref + "\" t=\"inlineStr\"");
            if (styleIndex != -1) sBuffer.append(" s=\"" + styleIndex + "\"");
            sBuffer.append(">");
            sBuffer.append("<is><t>" + value + "</t></is>");
            sBuffer.append("</c>");
        }

        public void createCell(int columnIndex, String value) throws IOException {
            createCell(columnIndex, value, -1);
        }

        public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
            String ref = new CellReference(_rownum, columnIndex).formatAsString();
            sBuffer.append("<c r=\"" + ref + "\" t=\"n\"");
            if (styleIndex != -1) sBuffer.append(" s=\"" + styleIndex + "\"");
            sBuffer.append(">");
            sBuffer.append("<v>" + value + "</v>");
            sBuffer.append("</c>");
        }

        public void createCell(int columnIndex, double value) throws IOException {
            createCell(columnIndex, value, -1);
        }

        public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
            createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
        }
    }

    private static String getColWidth(String xmlWidth, int[] arrWidth) {
        xmlWidth += "<cols>";
        int idxWidth = 0;
        for (int i1 = 0; i1 < arrWidth.length; i1++) {
            idxWidth += 1;
            String widthCellStr = String.valueOf(arrWidth[i1] * 1.25);
            xmlWidth += "<col min=\"" + idxWidth + "\" bestFit=\"1\" max=\"" + idxWidth + "\" width=\"" + widthCellStr + "\" customWidth=\"1\"/>";
        }
        xmlWidth += "</cols>";
        return xmlWidth;
    }

    private static short[] getRGB(String rgb) {
        final short[] ret = new short[3];
        for (int i = 0; i < 3; i++) {
            ret[i] = Short.parseShort(rgb.substring(i * 2, i * 2 + 2), 16);
        }
        return ret;
    }
}


