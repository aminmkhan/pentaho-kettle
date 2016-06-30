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

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.BeforeClass;
import org.junit.Test;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.utils.TestUtils;
import org.pentaho.di.core.KettleEnvironment;
import org.pentaho.di.core.RowMetaAndData;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.value.ValueMetaString;
import org.pentaho.di.trans.TransHopMeta;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.TransTestFactory;
import org.pentaho.di.trans.steps.excelinput.ExcelInputField;
import org.pentaho.di.trans.steps.excelinput.ExcelInputMeta;
import org.pentaho.di.trans.steps.excelinput.SpreadSheetType;


/**
 * @author Amin Khan
 */
public class ExcelWriterStep_StyleFormatTest {

  private ExcelWriterStep step;
  private ExcelWriterStepMeta meta;

  @BeforeClass
  public static void setUpEnv() throws KettleException {
    KettleEnvironment.init();
  }

  @Test
  public void test_style_format_Hssf() throws KettleException, IOException, Exception {
    createStepMeta("xls");
    List<RowMetaAndData> result = executeTrans("xls");

    assertNotNull( result );
    assertEquals( 3, result.size() );
    assertEquals( 6, result.get( 0 ).getRowMeta().size() );
  }

  @Test
  public void test_style_format_Xssf() throws KettleException, IOException, Exception {
    createStepMeta("xlsx");
    List<RowMetaAndData> result = executeTrans("xlsx");

    assertNotNull( result );
    assertEquals( 3, result.size() );
    assertEquals( 6, result.get( 0 ).getRowMeta().size() );
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

  private List<RowMetaAndData> getRowMetaAndData() {
    List<RowMetaAndData> rmd = new ArrayList<RowMetaAndData>();
    RowMeta rm = new RowMeta();
    rm.addValueMeta( new ValueMetaString( "col1" ) );
    rm.addValueMeta( new ValueMetaString( "col2" ) );
    rm.addValueMeta( new ValueMetaString( "col3" ) );
    rm.addValueMeta( new ValueMetaString( "col4" ) );
    rm.addValueMeta( new ValueMetaString( "col5" ) );
    rm.addValueMeta( new ValueMetaString( "col6" ) );
    rmd.add( new RowMetaAndData( rm,
            new Object[] { "1000.010101", "123456.654321", "9999.7777", "121212.4343434", "0", "-1021.32" } ) );
    rmd.add( new RowMetaAndData( rm,
            new Object[] { "1000", "-123456.6", "80808.777721", "13.4", "8989898e-10", "123e12" } ) );
    return rmd;
  }

  private List<RowMetaAndData> executeTrans( String filepath ) throws Exception {
    String stepName = "Excel Writer";

    TransMeta transMeta = TransTestFactory.generateTestTransformation( null, meta, stepName );
    List<RowMetaAndData> inputList = getRowMetaAndData();
    TransTestFactory.executeTestTransformation( transMeta, TransTestFactory.INJECTOR_STEPNAME, stepName,
            TransTestFactory.DUMMY_STEPNAME, inputList );

    try {
      Thread.sleep( 500 );
    } catch ( InterruptedException ignore ) {
      // Wait a bit to ensure that the output file is properly closed
    }

    // Now, check the result
    String checkStepName = "Excel Input";
    ExcelInputMeta excelInput = new ExcelInputMeta();
    excelInput.setFileName( new String[] { filepath } );
    excelInput.setFileMask( new String[] { "" } );
    excelInput.setExcludeFileMask( new String[] { "" } );
    excelInput.setFileRequired( new String[] { "N" } );
    excelInput.setIncludeSubFolders( new String[]{ "N" } );
    excelInput.setSpreadSheetType( SpreadSheetType.POI );
    excelInput.setSheetName( new String[] { "Sheet10" } );
    excelInput.setStartColumn( new int[] { 0 } );
    excelInput.setStartRow( new int[] { 0 } );
    excelInput.setStartsWithHeader( false ); // Ensures that we can check the header names

    ExcelInputField[] fields = new ExcelInputField[6];
    for ( int i = 0; i < 6; i++ ) {
      fields[i] = new ExcelInputField();
      fields[i].setName( "field" + ( i + 1 ) );
    }
    excelInput.setField( fields );

    transMeta = TransTestFactory.generateTestTransformation( null, excelInput, checkStepName );

    //Remove the Injector hop, as it's not needed for this transformation
    TransHopMeta injectHop = transMeta.findTransHop( transMeta.findStep( TransTestFactory.INJECTOR_STEPNAME ),
            transMeta.findStep( stepName ) );
    transMeta.removeTransHop( transMeta.indexOfTransHop( injectHop ) );

    inputList = new ArrayList<RowMetaAndData>();
    List<RowMetaAndData> result =
            TransTestFactory.executeTestTransformation( transMeta, TransTestFactory.INJECTOR_STEPNAME, stepName,
                    TransTestFactory.DUMMY_STEPNAME, inputList );

    return result;
  }
}
