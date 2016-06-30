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

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.pentaho.di.core.KettleEnvironment;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.logging.LoggingObjectInterface;
import org.pentaho.di.trans.steps.mock.StepMockHelper;


import junit.framework.Assert;
import static org.mockito.Matchers.any;
import static org.mockito.Mockito.*;

/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  private ExcelWriterStep step;
  private ExcelWriterStepData data;
  private ExcelWriterStepMeta meta;
  private static StepMockHelper<ExcelWriterStepMeta, ExcelWriterStepData> helper;

  @BeforeClass
  public static void setUpEnv() throws KettleException {
    KettleEnvironment.init();
    helper =
        new StepMockHelper<>( "ExcelWriterStep_StyleFormatTest",
                ExcelWriterStepMeta.class, ExcelWriterStepData.class );
    when( helper.logChannelInterfaceFactory.create( any(), any( LoggingObjectInterface.class ) ) ).thenReturn(
            helper.logChannelInterface );
    when( helper.trans.isRunning() ).thenReturn( true );
  }

  @Before
  public void setUp() throws Exception {
    step = new ExcelWriterStep(
            helper.stepMeta, helper.stepDataInterface, 0, helper.transMeta, helper.trans );

    step = spy( step );
    // ignoring to avoid useless errors in log
    doNothing().when( step ).prepareNextOutputFile();

    data = new ExcelWriterStepData();

    step.init( helper.initStepMetaInterface, data );
    Assert.assertEquals( "Step init error.", 0, step.getErrors() );
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
}
