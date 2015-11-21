import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.List;
import java.util.ArrayList;
import java.util.Date;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.POIXMLException;

import java.io.IOException;
import java.io.BufferedReader;
import java.io.UnsupportedEncodingException;
import java.io.InputStreamReader;

import com.ericsson.otp.erlang.OtpErlangAtom;
import com.ericsson.otp.erlang.OtpErlangDecodeException;
import com.ericsson.otp.erlang.OtpErlangException;
import com.ericsson.otp.erlang.OtpErlangExit;
import com.ericsson.otp.erlang.OtpErlangObject;
import com.ericsson.otp.erlang.OtpErlangPid;
import com.ericsson.otp.erlang.OtpErlangTuple;
import com.ericsson.otp.erlang.OtpErlangString;
import com.ericsson.otp.erlang.OtpErlangBinary;
import com.ericsson.otp.erlang.OtpErlangList;
import com.ericsson.otp.erlang.OtpErlangInt;
import com.ericsson.otp.erlang.OtpMbox;
import com.ericsson.otp.erlang.OtpNode;

public class XLSXReader {

  public static OtpNode	NODE;
  public static String	PEER;

  public static void main(String[] args) {
    String peerName = args[0];
    String nodeName = args[1];

    try {
      NODE = new OtpNode(nodeName, args[2]);
      PEER = peerName;

      final OtpMbox mbox = NODE.createMbox("xlsx_reader_server");

      new Thread(mbox.getName()) {
        @Override
        public void run() {
          boolean run = true;
          while (run) { // This thread runs forever
            try {
              run = processMsg(mbox.receive(), mbox);
            } catch (OtpErlangExit oee) {
              System.exit(1);
            } catch (OtpErlangDecodeException oede) {
              oede.printStackTrace();
              System.out.println("That was a bad message, moving on...");
            } catch (OtpErlangException oee) {
              oee.printStackTrace();
              System.out.println("That was a bad message, moving on...");
            }
          }
        }
      }.start();

      System.out.println("READY"); // Signal erlang we're running

      BufferedReader systemIn = new BufferedReader(new InputStreamReader(System.in, "UTF-8"));
      String line;
      while((line = systemIn.readLine()) != null) {
        Thread.sleep(1000);
      }
      System.exit(0); // Got SIGPIPE. Erlang closed Port, so exit
    } catch (Exception e1) {
      e1.printStackTrace();
      System.exit(1);
    }
  }

  protected static boolean processMsg(OtpErlangObject msg, OtpMbox mbox)
    throws OtpErlangDecodeException, OtpErlangException {
    if (msg instanceof OtpErlangAtom
        && ((OtpErlangAtom) msg).atomValue().equals("stop")) {
      mbox.close();
      return false;
    } else if (msg instanceof OtpErlangTuple) {
      OtpErlangObject[] elements = ((OtpErlangTuple) msg).elements();
      if (elements.length == 3) {
        if (elements[0] instanceof OtpErlangAtom
            && ((OtpErlangAtom) elements[0]).atomValue().equals(
                "parse") && elements[1] instanceof OtpErlangPid && elements[2] instanceof OtpErlangString) {
          OtpErlangPid caller = (OtpErlangPid) elements[1];
          String filename = ((OtpErlangString) elements[2]).stringValue();
          OtpErlangObject output = processFile(filename);
          mbox.send(caller, new OtpErlangTuple(new OtpErlangObject[] {
                new OtpErlangAtom("result"), mbox.self(), output}));
          return true;
        }
      }
    }
    throw new OtpErlangDecodeException("Bad message: " + msg.toString());
  }

  public static void padEmptyCells(List<OtpErlangObject> outputColls, int prevIndex, int index) {
    if (index > prevIndex + 1) {
      for(int i = 0; i < index-prevIndex-1; ++i)
        outputColls.add(new OtpErlangAtom("nil"));
    }
  }

  public static OtpErlangObject processFile(String filename) {
    FileInputStream fis = null;
    XSSFWorkbook book = null;

    try {
      File excel = new File(filename);
      fis = new FileInputStream(excel);
      book = new XSSFWorkbook(fis);
      XSSFSheet sheet = book.getSheetAt(0);
      HSSFDataFormatter formatter = new HSSFDataFormatter();

      Iterator<Row> itr = sheet.iterator();

      List<OtpErlangObject> outputRows = new ArrayList<>();

      while (itr.hasNext()) {
        boolean isRowEmpty = true;
        List<OtpErlangObject> outputColls = new ArrayList<>();

        Row row = itr.next();

        Iterator<Cell> cellIterator = row.cellIterator();
        int prevIndex = -1;
        int index = 0;

        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          int cellType = cell.getCellType();

          index = cell.getColumnIndex();

          if (cellType == Cell.CELL_TYPE_FORMULA) {
            cellType = cell.getCachedFormulaResultType();
          }

          switch (cellType) {
          case Cell.CELL_TYPE_STRING:
            isRowEmpty = false;
            padEmptyCells(outputColls, prevIndex, index);
            prevIndex = index;
            outputColls.add(convert(cell.getStringCellValue()));
            break;
          case Cell.CELL_TYPE_NUMERIC:
            isRowEmpty = false;
            padEmptyCells(outputColls, prevIndex, index);
            prevIndex = index;
            if (DateUtil.isCellDateFormatted(cell)) {
              outputColls.add(convert(cell.getDateCellValue()));
            } else {
              outputColls.add(convertNumber(formatter.formatCellValue(cell)));
            }
            break;
          default:
            break;
          }
        }

        if (!isRowEmpty) {
          OtpErlangObject[] array = outputColls.toArray(new OtpErlangObject[outputColls.size()]);
          outputRows.add(new OtpErlangList(array));
        }
      }

      OtpErlangObject[] array = outputRows.toArray(new OtpErlangObject[outputRows.size()]);
      return new OtpErlangTuple(new OtpErlangObject[] { new OtpErlangAtom("ok"), new OtpErlangList(array) });

    } catch (FileNotFoundException fe) {
      return new OtpErlangTuple(new OtpErlangObject[] { new OtpErlangAtom("error"), new OtpErlangAtom("enoent") });
    } catch (IOException ie) {
      return new OtpErlangTuple(new OtpErlangObject[] { new OtpErlangAtom("error"), new OtpErlangAtom("enoent") });
    } catch (POIXMLException pxe) {
      return new OtpErlangTuple(new OtpErlangObject[] { new OtpErlangAtom("error"), new OtpErlangAtom("format") });
    } finally {
      if (null != book) {
        try {
          book.close();
        } catch (IOException ioex) {
        }
      }
      if (null != fis) {
        try {
          fis.close();
        } catch (IOException ioex) {
        }
      }
    }
  }

  protected static OtpErlangObject convert(String string) {
    try {
      return new OtpErlangBinary(string.getBytes("UTF-8"));
    } catch(UnsupportedEncodingException uee) {
      return new OtpErlangAtom("nil");
    }
  }

  protected static OtpErlangObject convert(Date date) {
    Calendar cal = Calendar.getInstance();
    cal.setTime(date);
    int year = cal.get(Calendar.YEAR);
    int month = cal.get(Calendar.MONTH);
    int day = cal.get(Calendar.DAY_OF_MONTH);

    return new OtpErlangTuple(new OtpErlangObject[] {
        new OtpErlangTuple(new OtpErlangObject[] {new OtpErlangInt(year), new OtpErlangInt(month+1), new OtpErlangInt(day)}),
        new OtpErlangTuple(new OtpErlangObject[] {new OtpErlangInt(0), new OtpErlangInt(0), new OtpErlangInt(0)})
      });
  }

  protected static OtpErlangObject convertNumber(String string) {
    try {
      return new OtpErlangBinary(string.replace(",", "").getBytes("UTF-8"));
    } catch(UnsupportedEncodingException uee) {
      return new OtpErlangAtom("nil");
    }
  }

}
