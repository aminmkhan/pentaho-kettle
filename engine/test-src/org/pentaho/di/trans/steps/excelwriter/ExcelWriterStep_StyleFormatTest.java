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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.pentaho.di.core.KettleEnvironment;
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
  private static StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData> helper;

  @BeforeClass
  public static void setUpEnv() throws KettleException {
    KettleEnvironment.init();
    helper =
        new StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData>( "ExcelWriterStep_StyleFormatTest", ExcelWriterStepMeta.class,
                ExcelWriterStepData.class );
    when( helper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
            helper.logChannelInterface );
    when( helper.trans.isRunning() ).thenReturn( true );
  }

  @Before
  public void setUp() throws Exception {
    StepMockHelper<ExcelWriterStepMeta, StepDataInterface> mockHelper =
            StepMockUtil.getStepMockHelper( ExcelWriterStepMeta.class, "ExcelWriterStep_StyleFormatTest" );

    step = new ExcelWriterStep(
            mockHelper.stepMeta, mockHelper.stepDataInterface, 0, mockHelper.transMeta, mockHelper.trans );
    step = spy( step );
    // ignoring to avoid useless errors in log
    doNothing().when( step ).prepareNextOutputFile();

    data = new ExcelWriterStepData();

    step.init( mockHelper.initStepMetaInterface, data );
  }

  @Test
  public void generate_Hssf() throws Exception {
    data.wb = new HSSFWorkbook();
    data.wb.createSheet( "sheet1" );
    data.wb.createSheet( "sheet2" );

  }

  @Test
  public void generate_Xssf() throws Exception {
    data.wb = new XSSFWorkbook();
    data.wb.createSheet( "sheet1" );
    data.wb.createSheet( "sheet2" );

  }
}
