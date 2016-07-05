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

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.core.RowSet;
import org.pentaho.di.core.row.ValueMetaAndData;
import org.pentaho.di.core.row.value.ValueMetaBigNumber;
import org.pentaho.di.core.row.value.ValueMetaInteger;
import org.pentaho.di.core.row.value.ValueMetaNumber;
import org.pentaho.di.core.vfs.KettleVFS;
import org.pentaho.di.utils.TestUtils;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.logging.LoggingObjectInterface;
import org.pentaho.di.core.row.value.ValueMetaFactory;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.core.row.value.ValueMetaString;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.steps.mock.StepMockHelper;

import java.io.OutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.math.BigDecimal;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.*;
import static org.junit.Assert.assertEquals;
import static org.mockito.Matchers.*;
import static org.mockito.Mockito.*;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;


/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  private Workbook wb;
  private StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData> stepMockHelper;
  private ExcelWriterStep step;
  private ExcelWriterStepMeta stepMeta;
  private ExcelWriterStepData stepData;

  private String path;
  private List<Object[]> rows = new ArrayList<Object[]>();

  @Before
  public void setUp() throws Exception {
    stepMockHelper =
      new StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData>(
        "Excel Writer Style Format Test", ExcelWriterStepMeta.class, ExcelWriterStepData.class );
    when( stepMockHelper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
        stepMockHelper.logChannelInterface );
    verify( stepMockHelper.logChannelInterface, never() ).logError( anyString() );
    verify( stepMockHelper.logChannelInterface, never() ).logError( anyString(), any( Object[].class ) );
    verify( stepMockHelper.logChannelInterface, never() ).logError( anyString(), (Throwable) anyObject() );
    when( stepMockHelper.trans.isRunning() ).thenReturn( true );
  }

  @After
  public void tearDown() {
    stepMockHelper.cleanUp();
  }

  @Test
  public void testStyleFormatHssf() throws Exception {
    String fileType = "xls";
    setupStepMock( fileType );
    createStepMeta( fileType );
    createStepData( fileType );
    step.init( stepMeta, stepData );

    // TODO Try to load the template file present for ExcelOutput step
    String templateFilePath = "/org/pentaho/di/trans/steps/exceloutput/chart-template.xls";
    templateFilePath = "chart-template.xls";

    // step.prepareNextOutputFile();
    step.writeNextLine( rows.get(0) );

  }

  // @Test
  public void testStyleFormatXssf() throws Exception {
    // TODO Test for both xls and xlsx
    String fileType = "xlsx";
    setupStepMock( fileType );
    createStepMeta( fileType );
    createStepData( fileType );
  }

  // @Test
  public void testStyleFormatNoTemplate() throws Exception {
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );

    stepMeta.setTemplateEnabled( false );
    stepMeta.setTemplateFileName( "" );

    // step.prepareNextOutputFile();

    // List<RowMetaAndData> list = createRowData();



    // Tests

  }

  private void createStepMeta( String fileType ) throws KettleException, IOException {
    stepMeta = new ExcelWriterStepMeta();
    stepMeta.setDefault();

    stepMeta.setFileName( path.replace( "." + fileType, "" ) );
    stepMeta.setExtension( fileType );
    stepMeta.setSheetname( "Sheet101" );
    stepMeta.setHeaderEnabled( true );
    stepMeta.setStartingCell( "B3" );

    stepMeta.setTemplateEnabled( true );
    stepMeta.setTemplateFileName( "testExcelStyle." + fileType );
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
    stepData = new ExcelWriterStepData();

    // start writing from cell B3 in template
    stepData.startingRow = 2;
    stepData.startingCol = 1;
    stepData.posX = stepData.startingCol;
    stepData.posY = stepData.startingRow ;

    stepData.inputRowMeta = step.getInputRowMeta().clone();
    stepData.outputRowMeta = step.getInputRowMeta().clone();
    stepData.firstFileOpened = true;

    stepData.clearStyleCache( 4 );

    stepData.wb = stepMeta.getExtension().equalsIgnoreCase( fileType ) ? new XSSFWorkbook() : new HSSFWorkbook();
    stepData.sheet = stepData.wb.createSheet();
    // stepData.file = KettleVFS.getFileObject( buildFilename, getTransMeta() );

    stepData.fieldnrs = new int[] {0, 1, 2, 3};
  }

  private void setupStepMock( String fileType ) throws Exception {
    // TODO Avoid even creating file in RAM
    path = TestUtils.createRamFile( getClass().getSimpleName() + "/testExcelStyle." + fileType );
    FileObject xlsFile = TestUtils.getFileObject( path );
    wb = createWorkbook( xlsFile );

    step =
      new ExcelWriterStep(
        stepMockHelper.stepMeta, stepMockHelper.stepDataInterface, 0, stepMockHelper.transMeta, stepMockHelper.trans );
    step.init( stepMockHelper.initStepMetaInterface, stepMockHelper.initStepDataInterface );

    rows = createRowData();
    String[] outFields = new String[] { "col 1", "col 2", "col 3", "col 4" };
    RowSet inputRowSet = stepMockHelper.getMockInputRowSet( rows );
    RowMetaInterface inputRowMeta = createRowMeta();
    inputRowSet.setRowMeta( inputRowMeta );
    RowMetaInterface mockOutputRowMeta = mock( RowMetaInterface.class );
    when( mockOutputRowMeta.size() ).thenReturn( outFields.length );
    when( inputRowSet.getRowMeta() ).thenReturn( inputRowMeta );

    step.getInputRowSets().add( inputRowSet );
    step.setInputRowMeta( inputRowMeta );
    step.getOutputRowSets().add( inputRowSet );

    // step = spy( step );
    // // ignoring to avoid useless errors in log
    // doNothing().when( step ).prepareNextOutputFile();
    // TransMeta mockTransMeta = mock( TransMeta.class );
    // Trans mockTrans = mock( Trans.class );
    // when( step.getTransMeta() ).thenReturn( mockTransMeta );
    // when( step.getTrans() ).thenReturn( mockTrans );
  }

  private ArrayList<Object[]> createRowData() throws Exception {
    ArrayList<Object[]> r = new ArrayList<Object[]>();
    if( false ) {
      Object[] row = new Object[] {new Integer(1000), new Double(2.34e-4), new Double(40120), new Long(5010)};
      r.add(row);
      row = new Object[] {new Integer(123456), new Double(4.6789e10), new Double(111111e-2), new Long(12312300)};
      r.add(row);
    } else {
      // ValueMetaAndData, but may not work during writeLine()
      Object[] row = new Object[]{
        new ValueMetaAndData("col 1", 1000),
        new ValueMetaAndData("col 2", new Double(2.34e-4)),
        new ValueMetaAndData("col 3", new BigDecimal("123456789.987654321")),
        new ValueMetaAndData("col 4", new Long(5010))
      };
      r.add(row);
    }
    return r;
  }

  private RowMetaInterface createRowMeta() throws KettleException{
    RowMetaInterface rm = new RowMeta();
    try {
      ValueMetaInterface[] valuesMeta = {
        new ValueMetaInteger( "col 1" ),
        new ValueMetaNumber( "col 2" ),
        new ValueMetaBigNumber( "col 3" ),
        new ValueMetaNumber( "col 4" )
      };
      for ( int i = 0; i < valuesMeta.length; i++ ) {
        rm.addValueMeta( valuesMeta[i] );
      }
    } catch ( Exception ex ) {
      return null;
    }
    return rm;
  }

  private HSSFWorkbook createWorkbook( FileObject file ) throws Exception {
    HSSFWorkbook wb = null;
    OutputStream os = null;
    try {
      os = file.getContent().getOutputStream();
      wb = new HSSFWorkbook();
      wb.createSheet( "Sheet1" );
      wb.write( os );
    } finally {
      os.flush();
      os.close();
    }
    return wb;
  }

}
