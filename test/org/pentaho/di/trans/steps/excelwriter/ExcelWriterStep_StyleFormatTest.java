/*! ******************************************************************************
 *
 * Pentaho Data Integration
 *
 * Copyright (C) 2002-2016 by Pentaho : http://www.pentaho.com
 *
 *******************************************************************************
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 ******************************************************************************/

package org.pentaho.di.trans.steps.excelwriter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.BeforeClass;
import org.junit.Test;
import org.pentaho.di.core.KettleEnvironment;
import org.pentaho.di.core.RowMetaAndData;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.core.row.value.ValueMetaString;
import org.pentaho.di.trans.TransHopMeta;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.TransTestFactory;
import org.pentaho.di.trans.steps.excelinput.ExcelInputField;
import org.pentaho.di.trans.steps.excelinput.ExcelInputMeta;

/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  @BeforeClass
  public static void setUpEnv() throws KettleException {
    KettleEnvironment.init();
  }

  @Before
  public void setUp() throws Exception {
    String stepName = "Excel Writer";
    ExcelWriterStepMeta meta = new ExcelWriterStepMeta();
    meta.setDefault();

    File tempOutputFile = File.createTempFile( "testPDI11374", ".xlsx" );
    tempOutputFile.deleteOnExit();
    meta.setFileName( tempOutputFile.getAbsolutePath().replace( ".xlsx", "" ) );
    meta.setExtension( "xlsx" );
    meta.setSheetname( "Sheet10" );
    meta.setOutputFields( new ExcelWriterStepField[] {} );
    meta.setHeaderEnabled( true );
  }

  @Test
  public void generate_Hssf() throws Exception {
    setupMeta( "style-template.xls" );

    data.wb = new HSSFWorkbook();
    data.wb.createSheet( "sheet1" );
    data.wb.createSheet( "sheet2" );
    // assertTrue(1 == 12);
    System.out.println("Hello!");
  }

  @Test
  public void generate_Xssf() throws Exception {
    setupMeta( "style-template.xlsx" );

    data.wb = new XSSFWorkbook();
    data.wb.createSheet( "sheet1" );
    data.wb.createSheet( "sheet2" );
    // Assert.fail();

  }

  @After
  public void terminate() throws Exception {
    step.dispose( meta, helper.initStepDataInterface );
    Assert.assertEquals( "Step dispose error", 0, step.getErrors() );
  }

  private void setupMeta(String templateFileName) throws IOException {
    File tempFile = File.createTempFile( "PDI_excel_tmp", ".tmp" );
    tempFile.deleteOnExit();

  }

  private List<RowMetaAndData> getRowMetaAndData() {
    List<RowMetaAndData> rmd = new ArrayList<RowMetaAndData>();
    RowMeta rm = new RowMeta();
    rm.addValueMeta( new ValueMetaString( "col1" ) );
    rm.addValueMeta( new ValueMetaString( "col2" ) );
    rm.addValueMeta( new ValueMetaString( "col3" ) );
    rm.addValueMeta( new ValueMetaString( "col4" ) );
    rm.addValueMeta( new ValueMetaString( "col5" ) );
    rm.addValueMeta( new ValueMetaString( "col6" ) );
    rmd.add( new RowMetaAndData( rm, new Object[] { "1000.010101", "123456.654321", "9999.7777", "121212.4343434", "0", "-1021.32" } ) );
    rmd.add( new RowMetaAndData( rm, new Object[] { "1000", "-123456.6", "80808.777721", "13.4", "8989898e-10", "123e12" } ) );
    return rmd;
  }

}
