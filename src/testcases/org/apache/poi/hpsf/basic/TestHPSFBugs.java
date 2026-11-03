/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.hpsf.basic;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Date;
import java.util.Map;

import junit.framework.TestCase;

import org.apache.poi.POIDataSamples;
import org.apache.poi.POIDocument;
import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.HPSFPropertiesOnlyDocument;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.Section;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.util.HexRead;

/**
 * Tests various bugs have been fixed
 */
public final class TestHPSFBugs extends TestCase {
   private static final POIDataSamples _samples = POIDataSamples.getHPSFInstance();
    
   /**
    * Ensure that we can create a new HSSF Workbook,
    *  then add some properties to it, save +
    *  reload, and still access & change them.
    */
   public void test48832() throws Exception {
      HSSFWorkbook wb = new HSSFWorkbook();
      
      // Starts empty
      assertNull(wb.getDocumentSummaryInformation());
      assertNull(wb.getSummaryInformation());
      
      // Add new properties
      wb.createInformationProperties();
      
      assertNotNull(wb.getDocumentSummaryInformation());
      assertNotNull(wb.getSummaryInformation());
      
      // Set initial values
      wb.getSummaryInformation().setAuthor("Apache POI");
      wb.getSummaryInformation().setKeywords("Testing POI");
      wb.getSummaryInformation().setCreateDateTime(new Date(12345));
      
      wb.getDocumentSummaryInformation().setCompany("Apache");
      
      
      // Save and reload
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      wb.write(baos);
      ByteArrayInputStream bais = 
         new ByteArrayInputStream(baos.toByteArray());
      wb = new HSSFWorkbook(bais);
      
      
      // Ensure changes were taken
      assertNotNull(wb.getDocumentSummaryInformation());
      assertNotNull(wb.getSummaryInformation());
 
      assertEquals("Apache POI", wb.getSummaryInformation().getAuthor());
      assertEquals("Testing POI", wb.getSummaryInformation().getKeywords());
      assertEquals(12345, wb.getSummaryInformation().getCreateDateTime().getTime());
      assertEquals("Apache", wb.getDocumentSummaryInformation().getCompany());
      
      
      // Set some more, save + reload
      wb.getSummaryInformation().setComments("Resaved");
      
      baos = new ByteArrayOutputStream();
      wb.write(baos);
      bais = new ByteArrayInputStream(baos.toByteArray());
      wb = new HSSFWorkbook(bais);
      
      // Check again
      assertNotNull(wb.getDocumentSummaryInformation());
      assertNotNull(wb.getSummaryInformation());
 
      assertEquals("Apache POI", wb.getSummaryInformation().getAuthor());
      assertEquals("Testing POI", wb.getSummaryInformation().getKeywords());
      assertEquals("Resaved", wb.getSummaryInformation().getComments());
      assertEquals(12345, wb.getSummaryInformation().getCreateDateTime().getTime());
      assertEquals("Apache", wb.getDocumentSummaryInformation().getCompany());
   }
   
   /**
    * Some files seem to want the length and data to be on a 4-byte boundary,
    * and without that you'll hit an ArrayIndexOutOfBoundsException after
    * reading junk
    */
   public void test54233() throws Exception {
       DocumentInputStream dis;
       NPOIFSFileSystem fs = 
               new NPOIFSFileSystem(_samples.openResourceAsStream("TestNon4ByteBoundary.doc"));
       
       dis = fs.createDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
       SummaryInformation si = (SummaryInformation)PropertySetFactory.create(dis);
       
       dis = fs.createDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
       DocumentSummaryInformation dsi = (DocumentSummaryInformation)PropertySetFactory.create(dis);
      
       // Test
       assertEquals("Microsoft Word 10.0", si.getApplicationName());
       assertEquals("", si.getTitle());
       assertEquals("", si.getAuthor());
       assertEquals("Cour de Justice", dsi.getCompany());
       
       
       // Write out and read back, should still be valid
       POIDocument doc = new HPSFPropertiesOnlyDocument(fs);
       ByteArrayOutputStream baos = new ByteArrayOutputStream();
       doc.write(baos);
       ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
       doc = new HPSFPropertiesOnlyDocument(new NPOIFSFileSystem(bais));
       
       // Check properties are still there
       assertEquals("Microsoft Word 10.0", si.getApplicationName());
       assertEquals("", si.getTitle());
       assertEquals("", si.getAuthor());
       assertEquals("Cour de Justice", dsi.getCompany());
   }
   
   /**
    * CodePage Strings can be zero length
    */
   public void test56138() throws Exception {
       DocumentInputStream dis;
       NPOIFSFileSystem fs = 
               new NPOIFSFileSystem(_samples.openResourceAsStream("TestZeroLengthCodePage.mpp"));
       
       dis = fs.createDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
       SummaryInformation si = (SummaryInformation)PropertySetFactory.create(dis);
       
       dis = fs.createDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
       DocumentSummaryInformation dsi = (DocumentSummaryInformation)PropertySetFactory.create(dis);
       
       // Test
       assertEquals("MSProject", si.getApplicationName());
       assertEquals("project1", si.getTitle());
       assertEquals("Jon Iles", si.getAuthor());
       
       assertEquals("", dsi.getCompany());
       assertEquals(2, dsi.getSectionCount());
   }

   /**
    * Bug 52372 - OutOfMemoryError parsing a word file.
    * The offset of the custom properties section is padded/invalid.
    */
   public void test52372() throws Exception {
       NPOIFSFileSystem fs =
               new NPOIFSFileSystem(_samples.openResourceAsStream("TestBug52372.doc"));
       try {
           DocumentInputStream dis = fs.createDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
           SummaryInformation si = (SummaryInformation)PropertySetFactory.create(dis);
           dis.close();
           assertEquals("Microsoft Word 11.3", si.getApplicationName());

           dis = fs.createDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
           DocumentSummaryInformation dsi = (DocumentSummaryInformation)PropertySetFactory.create(dis);
           dis.close();
           assertEquals("Hewlett-Packard", dsi.getCompany());
           // the custom properties section, which has a padded offset, is readable
           assertEquals(2, dsi.getSectionCount());

           CustomProperties cps = dsi.getCustomProperties();
           assertNotNull(cps);
           assertEquals("TC101927549990", cps.get("_TemplateID"));
       } finally {
           fs.close();
       }
   }

   /**
    * A corrupt (huge) property count must not lead to an OutOfMemoryError
    */
   public void testCorruptPropertyCount() throws Exception {
       byte[] bytes = HexRead.readFromString(
           // byte order, format, os version, class id, section count (1)
           "FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 " +
           // section format id and offset (48)
           "E0 85 9F F2 F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00 " +
           // section size (24) and property count (0x7FFFFFFF)
           "18 00 00 00 FF FF FF 7F " +
           // property list: codepage at offset 16
           "01 00 00 00 10 00 00 00 " +
           // codepage property (VT_I2, 1252)
           "02 00 00 00 E4 04 00 00");
       try {
           new PropertySet(bytes);
           fail("A property count exceeding the stream should be detected");
       } catch (RuntimeException e) {
           // expected - and no OutOfMemoryError
       }
   }

   /**
    * A corrupt (huge) section count must not lead to an OutOfMemoryError
    */
   public void testCorruptSectionCount() throws Exception {
       byte[] bytes = HexRead.readFromString(
           // byte order, format, os version, class id, section count (0x7FFFFFFF)
           "FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF 7F");
       try {
           new PropertySet(bytes);
           fail("A section count exceeding the stream should be detected");
       } catch (RuntimeException e) {
           // expected - and no OutOfMemoryError
       }
   }

   /**
    * A corrupt (huge) number of dictionary entries must not lead to an OutOfMemoryError
    */
   public void testCorruptDictionaryEntryCount() throws Exception {
       byte[] bytes = HexRead.readFromString(
           // byte order, format, os version, class id, section count (1)
           "FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 " +
           // section format id and offset (48)
           "E0 85 9F F2 F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00 " +
           // section size (48) and property count (2)
           "30 00 00 00 02 00 00 00 " +
           // property list: codepage at offset 24, dictionary at offset 32
           "01 00 00 00 18 00 00 00 00 00 00 00 20 00 00 00 " +
           // codepage property (VT_I2, 1252)
           "02 00 00 00 E4 04 00 00 " +
           // dictionary with 0x7FFFFFFF entries, but only one entry (2 -> "Foo") present
           "FF FF FF 7F 02 00 00 00 04 00 00 00 46 6F 6F 00");
       PropertySet ps = new PropertySet(bytes);
       Section s = ps.getSections().get(0);
       assertEquals(1252, s.getCodepage());
       Map<Long,String> dic = s.getDictionary();
       assertNotNull(dic);
       assertEquals(1, dic.size());
       assertEquals("Foo", dic.get(2L));
   }

   /**
    * A corrupt (huge) dictionary string length must not lead to an OutOfMemoryError
    */
   public void testCorruptDictionaryStringLength() throws Exception {
       byte[] bytes = HexRead.readFromString(
           // byte order, format, os version, class id, section count (1)
           "FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 " +
           // section format id and offset (48)
           "E0 85 9F F2 F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00 " +
           // section size (48) and property count (2)
           "30 00 00 00 02 00 00 00 " +
           // property list: codepage at offset 24, dictionary at offset 32
           "01 00 00 00 18 00 00 00 00 00 00 00 20 00 00 00 " +
           // codepage property (VT_I2, 1200 / unicode)
           "02 00 00 00 B0 04 00 00 " +
           // dictionary with 1 entry (id 2), claiming 0x3FFFFFFF unicode characters
           "01 00 00 00 02 00 00 00 FF FF FF 3F 46 00 6F 00");
       PropertySet ps = new PropertySet(bytes);
       Section s = ps.getSections().get(0);
       assertEquals(1200, s.getCodepage());
       Map<Long,String> dic = s.getDictionary();
       assertNotNull(dic);
       assertTrue(dic.isEmpty());
   }
}
