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

import java.util.ArrayList;
import java.util.List;
import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Before;
import org.junit.Test;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.utils.TestUtils;
import org.pentaho.di.core.KettleEnvironment;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.logging.LoggingObjectInterface;
import org.pentaho.di.core.RowMetaAndData;
import org.pentaho.di.core.row.value.ValueMetaFactory;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.steps.StepMockUtil;
import org.pentaho.di.trans.steps.mock.StepMockHelper;

import static org.junit.Assert.*;
import static org.mockito.Mockito.*;


/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  private ExcelWriterStep step;
  private ExcelWriterStepData data;
  private ExcelWriterStepMeta meta;
  private static StepMockHelper<ExcelWriterStepMeta, StepDataInterface> helper;

  @BeforeClass
  public static void setUpEnv() throws KettleException {
//    KettleEnvironment.init( false );
  }

  @Before
  public void setUp() throws Exception {
//    helper =
//            StepMockUtil.getStepMockHelper( ExcelWriterStepMeta.class, "ExcelWriterStep_StyleFormatTest" );
//    when( helper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
//            helper.logChannelInterface );
//    when( helper.trans.isRunning() ).thenReturn( true );
//
//    step = new ExcelWriterStep(
//            helper.stepMeta, helper.stepDataInterface, 0, helper.transMeta, helper.trans );
//    step = spy( step );
//    // ignoring to avoid useless errors in log
//    doNothing().when( step ).prepareNextOutputFile();

    meta = new ExcelWriterStepMeta();
    meta.setDefault();
    createStepMeta( "xlsx" );
    data = new ExcelWriterStepData();

    step = new ExcelWriterStep( new StepMeta(), data, 0, new TransMeta(), new Trans() );
    step.init( meta, data );
  }

  @Test
  public void testStyleFormatHssf() throws Exception {
    createStepMeta( "xls" );
    createStepData( "xls" );

  }

  @Test
  public void testStyleFormatXssf() throws Exception {
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );
    step.init( meta, data );
    List<RowMetaAndData> list = createData();

    // step.prepareNextOutputFile();

    Object[] row = new Object[] {new Integer(1000), new Double(2.34e-4), new Double(40120), new Long(5010)};
    step.writeNextLine( row );

    // Tests
    Object v = null;
    for ( int i = 0; i < 4; i++ ) {
      v = row[i];
      System.out.println(v);
    }
  }

  @Test
  public void testStyleFormatNoTemplate() throws Exception {
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );

    meta.setTemplateEnabled( false );
    meta.setTemplateFileName( "" );

    // step.prepareNextOutputFile();

    // List<RowMetaAndData> list = createData();



    // Tests

  }

  private void createStepMeta( String fileType ) throws KettleException, IOException {
    // TODO Try to load the template file present for ExcelOutput step
    String templateFilePath = "/org/pentaho/di/trans/steps/exceloutput/chart-template.xls";
    templateFilePath = "chart-template.xls";
    meta.setDefault();

    // TODO Avoid even creating file in RAM
    String path = TestUtils.createRamFile( getClass().getSimpleName() + "/testExcelStyle." + fileType );
    FileObject xlsFile = TestUtils.getFileObject( path );

    meta.setFileName( path.replace( "." + fileType, "" ) );
    meta.setExtension( fileType );
    meta.setSheetname( "Sheet101" );
    meta.setHeaderEnabled( true );
    meta.setStartingCell( "B3" );

    meta.setTemplateEnabled( true );
    meta.setTemplateFileName( getClass().getResource( templateFilePath ).getFile() );
    meta.setTemplateSheetName( "SheetAsWell" );

    ExcelWriterStepField[] outputFields = new ExcelWriterStepField[4];
    outputFields[0] = new ExcelWriterStepField( "col 1", ValueMetaFactory.getIdForValueMeta( "Number" ), "" );
    outputFields[0].setStyleCell( "B3" );
    outputFields[1] = new ExcelWriterStepField( "col 2", ValueMetaFactory.getIdForValueMeta( "BigNumber" ), "0" );
    outputFields[1].setStyleCell( "B2" );
    outputFields[2] = new ExcelWriterStepField( "col 3", ValueMetaFactory.getIdForValueMeta( "Integer" ), "0.0" );
    outputFields[2].setStyleCell( "" );
    outputFields[3] = new ExcelWriterStepField( "col 4", ValueMetaFactory.getIdForValueMeta( "Integer" ), "0.0000" );
    outputFields[3].setStyleCell( "B2" );

    meta.setOutputFields( outputFields );

  }

  private void createStepData( String fileType ) throws KettleException {

    // start writing from cell B3 in template
    data.startingRow = 2;
    data.startingCol = 1;
    data.posX = data.startingCol;
    data.posY = data.startingRow ;

    data.outputRowMeta = new RowMeta();
    data.inputRowMeta = new RowMeta();
    data.firstFileOpened = true;

    data.wb = meta.getExtension().equalsIgnoreCase( fileType ) ? new XSSFWorkbook() : new HSSFWorkbook();
    data.sheet = data.wb.createSheet();

    data.inputRowMeta = null;
  }

  private List<RowMetaAndData> createData() throws KettleException {
    List<RowMetaAndData> list = new ArrayList<RowMetaAndData>();

    return list;
  }

}
