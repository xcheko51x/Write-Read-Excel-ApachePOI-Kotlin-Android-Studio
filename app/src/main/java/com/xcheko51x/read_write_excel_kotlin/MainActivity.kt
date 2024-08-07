package com.xcheko51x.read_write_excel_kotlin

import android.Manifest
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import com.xcheko51x.read_write_excel_kotlin.databinding.ActivityMainBinding
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding
    private lateinit var adapter: UsuarioAdapter

    var listaRegistros = listOf<Usuario>()

    private lateinit var solicitarPermisos: ActivityResultLauncher<Array<String>>

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        solicitarPermisos = registerForActivityResult(ActivityResultContracts.RequestMultiplePermissions()) {
            val aceptados = it.all { it.value }
            if (aceptados) {
                // OPERACIONES PERMITIDAS
            } else {
                Toast.makeText(this, "SE TIENE QUE ACEPTAR TODOS LOS PERMISOS", Toast.LENGTH_SHORT).show()
            }
        }

        permisos()

        setupRecyclerView()

        binding.btnAgregar.setOnClickListener {
            val lista = listaRegistros.toMutableList()
            lista.add(
                Usuario(
                    binding.etNombre.text.toString(),
                    binding.etEdad.text.toString()
                )
            )
            listaRegistros = lista

            binding.etNombre.setText("")
            binding.etEdad.setText("")

            setupRecyclerView()
        }

        binding.btnEscribir.setOnClickListener {
            if (listaRegistros.isNotEmpty()) {
                crearExcel(listaRegistros.toMutableList())
            }
        }

        binding.btnLeer.setOnClickListener {
            listaRegistros = leerExcel(listaRegistros.toMutableList())
            setupRecyclerView()
        }

    }

    fun setupRecyclerView() {
        adapter = UsuarioAdapter(listaRegistros)
        binding.rvListaRegistros.adapter = adapter
    }

    fun permisos() {
        solicitarPermisos.launch(
            arrayOf(
                Manifest.permission.WRITE_EXTERNAL_STORAGE,
                Manifest.permission.READ_EXTERNAL_STORAGE
            )
        )
    }

    fun crearExcel(listaRegistros: MutableList<Usuario>) {
        val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)

        val fileName = "registros.xlsx"

        // Crear un nuevo libro de trabajo Excel en formato .xlsx
        val workbook = XSSFWorkbook()

        // Crear una hoja de trabajo (worksheet)
        val sheet: Sheet = workbook.createSheet("Hoja 1")

        // Crear una fila en la hoja
        val headerRow = sheet.createRow(0)

        // Crear celdas en la fila
        var cell = headerRow.createCell(0)
        cell.setCellValue("Nombre")

        cell = headerRow.createCell(1)
        cell.setCellValue("Edad")

        for (index in listaRegistros.indices) {
            val row = sheet.createRow(index + 1)
            row.createCell(0).setCellValue(listaRegistros[index].nombre)
            row.createCell(1).setCellValue(listaRegistros[index].edad)
        }

        // Guardar el libro de trabajo (workbook) en almacenamiento externo
        try {
            val fileOutputStream = FileOutputStream(
                File(path, fileName)
            )
            workbook.write(fileOutputStream)
            fileOutputStream.close()
            workbook.close()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    fun leerExcel(listaRegistros: MutableList<Usuario>): MutableList<Usuario> {
        val fileName = "registros.xlsx"

        val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS).absolutePath+"/"+fileName

        Log.d("EXCEL","EXCEL LEER")

        val lista = arrayListOf<String>()

        try {
            val fileInputStream = FileInputStream(path)
            val workbook = WorkbookFactory.create(fileInputStream)
            val sheet: Sheet = workbook.getSheetAt(0)

            val rows = sheet.iterator()
            while (rows.hasNext()) {
                val currentRow = rows.next()

                // Iterar sobre celdas de la fila actual
                val cellsInRow = currentRow.iterator()
                while (cellsInRow.hasNext()) {
                    val currentCell = cellsInRow.next()

                    // Obtener valor de la celda como String
                    val cellValue: String = when (currentCell.cellType) {
                        CellType.STRING -> currentCell.stringCellValue
                        CellType.NUMERIC -> currentCell.numericCellValue.toString()
                        CellType.BOOLEAN -> currentCell.booleanCellValue.toString()
                        else -> ""
                    }

                    lista.add(cellValue)
                }
            }

            for (i in 2 until lista.size step 2) {
                listaRegistros.add(
                    Usuario(
                        lista[i],
                        lista[i+1]
                    )
                )
            }

            workbook.close()
            fileInputStream.close()

        } catch (e: IOException) {
            e.printStackTrace()
        }

        return listaRegistros
    }
}