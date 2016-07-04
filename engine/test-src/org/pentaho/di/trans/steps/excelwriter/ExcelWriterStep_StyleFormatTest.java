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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Before;
import org.junit.Test;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.utils.TestUtils;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.logging.LoggingObjectInterface;
import org.pentaho.di.core.RowMetaAndData;
import org.pentaho.di.core.row.value.ValueMetaFactory;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.trans.steps.StepMockUtil;
import org.pentaho.di.trans.steps.mock.StepMockHelper;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;

import static org.junit.Assert.*;
import static org.junit.Assert.assertEquals;
import static org.mockito.Matchers.any;
import static org.mockito.Mockito.when;


/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  private static final String SHEET_NAME = "Sheet1";
  private static final String FILE_TYPE = "xls";

  private HSSFWorkbook wb;
  private StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData> mockHelper;
  private ExcelWriterStep step;
  private ExcelWriterStepMeta stepMeta;
  private ExcelWriterStepData stepData;

  private String path;

  @Before
  public void setUp() throws Exception {
    // TODO Avoid even creating file in RAM
    path = TestUtils.createRamFile( getClass().getSimpleName() + "/testExcelStyle." + FILE_TYPE );
    FileObject xlsFile = TestUtils.getFileObject( path );
    wb = createWorkbook( xlsFile );

    mockHelper =
      new StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData>(
        "Excel Writer Style Format Test", ExcelWriterStepMeta.class, ExcelWriterStepData.class );
    when( mockHelper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
        mockHelper.logChannelInterface );
    step =
      new ExcelWriterStep(
        mockHelper.stepMeta, mockHelper.stepDataInterface, 0, mockHelper.transMeta, mockHelper.trans );



    // TODO Do we need spy? Or do nothing when next output file?
    // step = spy( step );
    // // ignoring to avoid useless errors in log
    // doNothing().when( step ).prepareNextOutputFile();

    stepMeta = new ExcelWriterStepMeta();
    stepData = new ExcelWriterStepData();

    // TODO Do we need to initialize step?
    step.init( mockHelper.initStepMetaInterface, stepData );
  }

  @Test
  public void testStyleFormatHssf() throws Exception {
    createStepMeta( "xls" );
    createStepData( "xls" );


    // step.prepareNextOutputFile();



    // TODO Some redundant tests just to verify it is working
    step.protectSheet( wb.getSheet( SHEET_NAME ), "aa" );
    assertTrue( wb.getSheet( SHEET_NAME ).getProtect() );

  }

  // @Test
  public void testStyleFormatXssf() throws Exception {
    // TODO Test for both xls and xlsx
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );
  }

  // @Test
  public void testStyleFormatNoTemplate() throws Exception {
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );

    stepMeta.setTemplateEnabled( false );
    stepMeta.setTemplateFileName( "" );

    // step.prepareNextOutputFile();

    // List<RowMetaAndData> list = createData();



    // Tests

  }

  private void createStepMeta( String fileType ) throws KettleException, IOException {
    // TODO Try to load the template file present for ExcelOutput step
    String templateFilePath = "/org/pentaho/di/trans/steps/exceloutput/chart-template.xls";
    templateFilePath = "chart-template.xls";
    stepMeta.setDefault();

    stepMeta.setFileName( path.replace( "." + fileType, "" ) );
    stepMeta.setExtension( fileType );
    stepMeta.setSheetname( "Sheet101" );
    stepMeta.setHeaderEnabled( true );
    stepMeta.setStartingCell( "B3" );

    stepMeta.setTemplateEnabled( true );
    stepMeta.setTemplateFileName( getClass().getResource( templateFilePath ).getFile() );
    stepMeta.setTemplateSheetName( "SheetAsWell" );

    ExcelWriterStepField[] outputFields = new ExcelWriterStepField[4];
    outputFields[0] = new ExcelWriterStepField( "col 1", ValueMetaFactory.getIdForValueMeta( "Number" ), "" );
    outputFields[0].setStyleCell( "B3" );
    outputFields[1] = new ExcelWriterStepField( "col 2", ValueMetaFactory.getIdForValueMeta( "BigNumber" ), "0" );
    outputFields[1].setStyleCell( "B2" );
    outputFields[2] = new ExcelWriterStepField( "col 3", ValueMetaFactory.getIdForValueMeta( "Integer" ), "0.0" );
    outputFields[2].setStyleCell( "" );
    outputFields[3] = new ExcelWriterStepField( "col 4", ValueMetaFactory.getIdForValueMeta( "Integer" ), "0.0000" );
    outputFields[3].setStyleCell( "B2" );

    stepMeta.setOutputFields( outputFields );

  }

  private void createStepData( String fileType ) throws KettleException {

    // start writing from cell B3 in template
    stepData.startingRow = 2;
    stepData.startingCol = 1;
    stepData.posX = stepData.startingCol;
    stepData.posY = stepData.startingRow ;

    stepData.outputRowMeta = new RowMeta();
    stepData.inputRowMeta = new RowMeta();
    stepData.firstFileOpened = true;

    stepData.wb = stepMeta.getExtension().equalsIgnoreCase( fileType ) ? new XSSFWorkbook() : new HSSFWorkbook();
    stepData.sheet = stepData.wb.createSheet();

    stepData.inputRowMeta = null;
  }

  private Object[] createData() throws Exception {
    Object[] row = new Object[] {new Integer(1000), new Double(2.34e-4), new Double(40120), new Long(5010)};
    return row;
  }

  private HSSFWorkbook createWorkbook( FileObject file ) throws Exception {
    HSSFWorkbook wb = null;
    OutputStream os = null;
    try {
      os = file.getContent().getOutputStream();
      wb = new HSSFWorkbook();
      wb.createSheet( SHEET_NAME );
      wb.write( os );
    } finally {
      os.flush();
      os.close();
    }
    return wb;
  }

}
