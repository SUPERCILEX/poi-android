package com.supercilex.poiandroid.demo

import android.Manifest
import android.content.Intent
import android.os.Bundle
import android.os.Environment
import androidx.core.content.FileProvider
import androidx.appcompat.app.AppCompatActivity
import android.util.Log
import android.view.View
import android.widget.Toast
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.charts.AxisCrosses
import org.apache.poi.ss.usermodel.charts.AxisPosition
import org.apache.poi.ss.usermodel.charts.DataSources
import org.apache.poi.ss.usermodel.charts.LegendPosition
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFChart
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle
import pub.devrel.easypermissions.EasyPermissions
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        setSupportActionBar(findViewById(R.id.toolbar))

        findViewById<View>(R.id.start).setOnClickListener {
            if (!EasyPermissions.hasPermissions(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
                EasyPermissions.requestPermissions(
                        this,
                        "Needed for the demo",
                        2,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE
                )
                return@setOnClickListener
            }

            async { generateExportIntent() }.addOnSuccessListener {
                startActivity(it)
            }.addOnFailureListener {
                Log.e("SpreadsheetExporter", "Export failed", it)
                Toast.makeText(this, it.stackTrace.toList().toString(), Toast.LENGTH_LONG).show()
            }
        }
    }

    override fun onRequestPermissionsResult(
            requestCode: Int,
            permissions: Array<out String>,
            grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        findViewById<View>(R.id.start).performClick()
    }

    private fun generateExportIntent(): Intent {
        val file = File(
                "${Environment.getExternalStorageDirectory()}/Documents/poi-android-demo.xlsx")

        if (file.exists()) file.delete()

        val workbook: Workbook = XSSFWorkbook()
        workbook.createSheet("Test").populate()
        workbook.write(FileOutputStream(file))

        val uri = FileProvider.getUriForFile(this, "$packageName.provider", file)
        return Intent()
                .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
                .setAction(Intent.ACTION_VIEW)
                .setDataAndType(uri, "application/vnd.ms-excel")
                .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                .also {
                    if (it.resolveActivity(packageManager) == null) {
                        it.setDataAndType(uri, "*/*")
                    }
                }
    }

    /**
     * Add random stuff to a spreadsheet to make sure everything works
     */
    private fun Sheet.populate() {
        createFreezePane(1, 1)
        createRow(0).apply {
            createCell(0).apply {
                setCellValue(0.0)
                cellStyle = workbook.createCellStyle().apply {
                    dataFormat = workbook.createDataFormat().getFormat("0.00")
                    setFont(workbook.createFont().apply { bold = true })
                    setAlignment(HorizontalAlignment.CENTER)
                    setVerticalAlignment(VerticalAlignment.CENTER)
                }
            }
            createCell(1).apply {
                setCellValue("Wassup?")
            }
            createCell(2).apply {
                createRow(1).createCell(0).setCellValue(1.0)
                cellFormula = "SUM(${CellRangeAddress(0, 1, 0, 0).formatAsString()})"
            }
        }

        createDrawingPatriarch().run {
            createChart(createAnchor(0, 0, 0, 0, 4, 2, 10, 6))
        }.apply {
            orCreateLegend.position = LegendPosition.RIGHT

            chartDataFactory.createLineChartData().apply {
                addSeries(
                        DataSources.fromStringCellRange(
                                this@populate,
                                CellRangeAddress(0, 0, 1, 1)
                        ),
                        DataSources.fromNumericCellRange(
                                this@populate,
                                CellRangeAddress(0, 1, 0, 0)
                        )
                ).setTitle("Foobar")
                plot(
                        this,
                        chartAxisFactory.createCategoryAxis(AxisPosition.BOTTOM),
                        chartAxisFactory.createValueAxis(AxisPosition.LEFT).apply {
                            crosses = AxisCrosses.AUTO_ZERO
                        }
                )
            }

            if (this is XSSFChart) {
                val plotArea = ctChart.plotArea
                plotArea.getValAxArray(0).addNewTitle().setValue("Values")
                plotArea.getCatAxArray(0).addNewTitle().setValue("Title")
            }
        }
    }

    private fun CTTitle.setValue(text: String) {
        addNewLayout()
        addNewOverlay().`val` = false

        val textBody = addNewTx().addNewRich()
        textBody.addNewBodyPr()
        textBody.addNewLstStyle()

        val paragraph = textBody.addNewP()
        paragraph.addNewPPr().addNewDefRPr()
        paragraph.addNewR().t = text
        paragraph.addNewEndParaRPr()
    }

    companion object {
        init {
            System.setProperty(
                    "org.apache.poi.javax.xml.stream.XMLInputFactory",
                    "com.fasterxml.aalto.stax.InputFactoryImpl"
            )
            System.setProperty(
                    "org.apache.poi.javax.xml.stream.XMLOutputFactory",
                    "com.fasterxml.aalto.stax.OutputFactoryImpl"
            )
            System.setProperty(
                    "org.apache.poi.javax.xml.stream.XMLEventFactory",
                    "com.fasterxml.aalto.stax.EventFactoryImpl"
            )
        }
    }
}
