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

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.vfs2.FileObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Before;
import org.junit.Test;
import org.pentaho.di.utils.TestUtils;
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
  public void test_style_format_Hssf() throws KettleException, IOException, Exception {
    createStepMeta("xls");

  }

  @Test
  public void test_style_format_Xssf() throws KettleException, IOException, Exception {
    createStepMeta("xlsx");

  }

  private void createStepMeta(String filetype) throws Exception {
    meta = new ExcelWriterStepMeta();
    meta.setDefault();

    String path = TestUtils.createRamFile( getClass().getSimpleName() + "/testExcelStyle." + filetype );
    FileObject xlsFile = TestUtils.getFileObject( path );

    meta.setFileName( path.replace( "." + filetype, "" ) );
    meta.setExtension( filetype );
    meta.setOutputFields( new ExcelWriterStepField[] {} );
    meta.setHeaderEnabled( true );
  }
}
