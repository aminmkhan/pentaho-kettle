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
import org.junit.Before;
import org.junit.Test;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.utils.TestUtils;
import org.pentaho.di.core.RowMetaAndData;
import org.pentaho.di.core.row.value.ValueMetaFactory;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.logging.LoggingObjectInterface;
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

  @Before
  public void setUp() throws Exception {
    helper =
            StepMockUtil.getStepMockHelper( ExcelWriterStepMeta.class, "ExcelWriterStep_StyleFormatTest" );
    when( helper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
            helper.logChannelInterface );
    when( helper.trans.isRunning() ).thenReturn( true );

    step = new ExcelWriterStep(
            helper.stepMeta, helper.stepDataInterface, 0, helper.transMeta, helper.trans );
    step = spy( step );
    // ignoring to avoid useless errors in log
    doNothing().when( step ).prepareNextOutputFile();

    data = new ExcelWriterStepData();
    meta = new ExcelWriterStepMeta();

    step.init( helper.initStepMetaInterface, data );
  }

  @Test
  public void test_style_format_Hssf() throws Exception {
    createStepMeta( "xls" );
    createStepData( "xls" );

  }

  @Test
  public void test_style_format_Xssf() throws Exception {
    createStepMeta( "xlsx" );
    createStepData( "xlsx" );
    List<RowMetaAndData> list = createData();

    // step.writeNextLine( list[0] );

    // Tests

  }

  private void createStepMeta( String filetype ) throws KettleException, IOException {
    // TODO Try to load the template file present for ExcelOutput step
    String templateFilePath = "/org/pentaho/di/trans/steps/exceloutput/chart-template.xls";
    templateFilePath = "chart-template.xls";
    meta.setDefault();

    // TODO Avoid even creating file in RAM
    String path = TestUtils.createRamFile( getClass().getSimpleName() + "/testExcelStyle." + filetype );
    FileObject xlsFile = TestUtils.getFileObject( path );

    meta.setFileName( path.replace( "." + filetype, "" ) );
    meta.setExtension( filetype );
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

  private void createStepData( String filetype ) throws KettleException {
    data.posX = 0;
    data.posY = 0;

    data.wb = null;
    data.sheet = null;

    data.inputRowMeta = null;
  }

  private List<RowMetaAndData> createData() throws KettleException {
    List<RowMetaAndData> list = new ArrayList<RowMetaAndData>();

    return list;
  }

}
