import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Scanner;
import java.util.regex.*;
import java.util.*;
import java.util.stream.*;
import java.io.*;
import java.util.concurrent.TimeUnit;
/*Estudiante Anderson Morillo Diaz*/

public class NotaDef {

    double quiz1,quiz2,quiz3,taller1,taller2;
    double nota1, nota2, nota3;
    String nombre,estado;
    Scanner entrada = new Scanner(System.in);
    String reportes[];
    public static final String SAMPLE_XLSX_FILE_PATH = "D:\\Anderson\\Downloads\\DATOS.xlsx";
    public static void main(String[] args) throws IOException, InvalidFormatException {
            NotaDef fc = new NotaDef();
            Estudiante arr[] = new Estudiante[0];
            String menu;
            Map <Integer, Estudiante> data= new HashMap<>();
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }
        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        Iterator<Row> rowIterator = sheet.rowIterator();
        int numEstudiantes=sheet.getLastRowNum();
        arr= new Estudiante[numEstudiantes];
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        int countRow=0,countRow1=0;
        Sheet sheet2 =workbook.createSheet("Reportes");
        for (Row row: sheet) {
            if (row.getRowNum()!=sheet.getFirstRowNum()){
                int countCell=0;
                for(Cell cell: row) {
                    boolean respuesta = getCellValue(cell,countCell,sheet2,countRow,sheet);
                    try{
                        if (!respuesta){
                            Row row1 = sheet2.createRow(countRow1++);
                            Cell cell2 = row1.createCell(0);
                            Cell name=row.getCell(1);
                            try{
                                cell2.setCellValue("Se cambio la nota en la fila "+cell.getAddress()+" del estudiante "+name.getNumericCellValue()+" debido a inconsistencia en el valor previamente establecido");//contador de notas
                            }catch (Exception e){
                                cell2.setCellValue("Se cambio la nota en la fila "+cell.getAddress()+" del estudiante "+name.getStringCellValue()+" debido a inconsistencia en el valor previamente establecido");
                            }
                        }
                    }catch (Exception e){

                    }

                    switch (cell.getColumnIndex()){
                        case 1:
                            try{
                                fc.nombre = cell.getStringCellValue();
                            }catch (Exception e){
                                fc.nombre = "Sin nombre";
                            }break;
                        case 2:
                            try{
                                fc.nota1 = cell.getNumericCellValue();
                            } catch (Exception e){
                                fc.nota1 = 0;
                            }break;
                        case 3:
                            try {
                                fc.quiz1 = cell.getNumericCellValue();
                            } catch (Exception e){
                                fc.quiz1 =0;
                            }break;
                        case 4:
                            try {
                                fc.quiz2 = cell.getNumericCellValue();
                            } catch (Exception e){
                                fc.quiz2 =0;
                            }break;
                        case 5:
                            try {
                                fc.quiz3 = cell.getNumericCellValue();
                            } catch (Exception e){
                                fc.quiz3 =0;
                            }break;
                        case 6:
                            try {
                                fc.taller1 = cell.getNumericCellValue();
                            } catch (Exception e) {
                                fc.taller1 = 0;
                            }break;
                        case 7:
                            try {
                                fc.taller2 =cell.getNumericCellValue();
                            } catch (Exception e){
                                fc.taller2=0;
                            }break;
                        default:
                            System.out.println(" ");
                    }
                }
            }
            if (row.getRowNum()!=sheet.getFirstRowNum()){
                try{
                    arr[countRow] = new Estudiante(fc.nombre,fc.nota1, fc.quiz1, fc.quiz2, fc.quiz3, fc.taller1,fc.taller2);
                    countRow++;
                } catch (Exception e){
                    break;
                }
                System.out.println(" ");
            }
        }
        int i=0;
        int contadorFilas=0;
        Cell cell1,cell2;
        boolean condicion=true;
        for (Row row: sheet){
            if (row.getRowNum()==sheet.getLastRowNum()+1){condicion=false;}
                if (row.getRowNum()!=0 && condicion){
                    cell1=row.createCell(8);
                    cell2=row.createCell(9);
                    cell1=row.getCell(8);
                    cell1.setCellValue(arr[contadorFilas].getNota());
                    cell2=row.getCell(9);
                    cell2.setCellValue(arr[contadorFilas].getEstado());
                    i++;
                    contadorFilas++;
                }else {
                    cell1=row.createCell(8);
                    cell2=row.createCell(9);
                    cell1.setCellValue("Promedio Nota");
                    cell2.setCellValue("Estado");
                }
        }
        Sheet sheet1=workbook.createSheet("Informe");
        fc.principal(arr.length,arr,sheet1);//arreglar retornando la hoja de calculo
        FileOutputStream out = new FileOutputStream(
                "D:/Anderson/Downloads/DATOS1.xlsx");
        workbook.write(out);
        out.close();
    }


        private static boolean getCellValue(Cell cell, int conCell,Sheet sheet2,int contador,Sheet sheet) {
        final double cero =0;
        double numero=0;
        int countRow=contador;
        Row row;
        Cell cell1;
        Row row1=sheet.getRow(contador);
        Cell name=row1.getCell(1);
        switch (cell.getCellType()) {
            case BOOLEAN:
            case FORMULA:
                //row=sheet2.createRow(countRow);
                //cell1=row.createCell(0);
                //name=row1.getCell(1);
                //cell1.setCellValue("Se cambio la nota en la fila "+cell.getRowIndex()+" nombre "+name.getStringCellValue()+" debido a inconsistencia en el valor previamente establecido");
                cell.setCellValue(cero);
                return false;
            case STRING:
                double res= numberValidation(cell.getStringCellValue());
                if (conCell>=2) {
                    cell.setCellValue(cero);
                } else {
                    if (res!=-1){
                        cell.setCellValue(res);
                        System.out.println(cell.getNumericCellValue());
                        return false;
                    } else {
                        String word = cell.getRichStringCellValue().getString();
                        if (word.length() < 4) {
                            cell.setCellValue(cero);
                            System.out.print(cell.getNumericCellValue());
                            return false;
                        } else {
                            System.out.print(cell.getStringCellValue());
                            return true;
                        }
                    }
                }
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.setCellValue(cero);
                    return false;
                } else {
                    numero=cell.getNumericCellValue();
                    if (numero>5 || numero<0) {
                        cell.setCellValue(cero);
                        return false;
                    }else {
                        cell.setCellValue(numero);
                        return true;
                    }
                }
            case BLANK:
                cell.setCellValue(cero);
                System.out.println(cell.getNumericCellValue());
                return  false;
            default:
                System.out.print("print default\n ");
        }
        System.out.print("\t");
        return true;
    }

    static double numberValidation(String prueba) {
        String noComma = prueba.replaceAll(",",".");
        boolean resp=isNumeric(noComma);
        if (resp){
            double number= Double.parseDouble(noComma);
            if (number<=5 && number>=0) {
                return  number;
            } else {
                return -1;
            }
        }
            return -1;
    }

    public void IngreseNotas(int i) {
        /*Metodo se encarga de ingresar la informacion del estudiante (nombre y notas de parcial, quices y talleres) */
        System.out.println("----------------------------informacion del estudiante "+i+ " --------------------------");
        System.out.println("Por favor ingrese informacion del estudiante " + i + " :");
        do {
            System.out.print("Ingresar nombre:");
            nombre = entrada.next();
            if (!isValidUsername(nombre)){
                System.out.println("Error al ingresar el nombre del estudiante!, ingresar nombre que comience con letra y tenga mas de 1 digito");
            }
        } while (!isValidUsername(nombre));
        System.out.println("introducir notas con valores entre 0 y 5:");
        nota1=IntroducirNota("Parcial");
        nota2=IntroducirNota("Quices");
        nota3=IntroducirNota("Talleres");
    }

    public double IntroducirNota(String nombreExamen){
        double nota=-1;
        do{
            System.out.print("Ingrese nota "+nombreExamen+":");
            String notaStr = entrada.next();
            boolean resp =isNumeric(notaStr);
            if (resp) {
                nota= Double.parseDouble(notaStr);
                if (nota > 5 || nota < 0) {
                    System.out.println("Error al digitar la nota introducida!, Digite numero dentro del rango 0 al 5:");
                    nota=IntroducirNota(nombreExamen);
                    return nota;
                }else {
                    return nota;
                }
            }else {
                System.out.println("Error al digitar la nota introducida!, no introdir texto.");
            }
        }while (nota > 5 || nota < 0);
        return nota;
    }

    private static boolean isNumeric(String cadena){
        try {
            double numero=Double.parseDouble(cadena);
            return true;
        } catch (NumberFormatException nfe){
            return false;
        }
    }

    private static int isNumericInt(String cadena){
        try {
            int numero=Integer.parseInt(cadena);
            return numero;
        } catch (NumberFormatException nfe){
            return -1;
        }
    }

    public int Mensaje(int i,Estudiante arr[]) {
        /*El metodo imprime si los estudiantes aprobaron o reprobaron y devuelve 1 o 0 para contabilizar */
        System.out.println("\n");
        System.out.println("-----------------------------------------notas---------------------------------------------");
        if (arr[i].getNota() >= 3 && arr[i].getNota() <= 5) {
            System.out.println("Estudiante "+arr[i].getNombre()+" nota: "+arr[i].getNota());
            System.out.println(" Aprobado");
            return 1;
        } else {
            if (arr[i].getNota() >= 0 && arr[i].getNota() < 3) {
                System.out.println("Estudiante "+arr[i].getNombre()+" y nota :"+arr[i].getNota());
                System.out.println("Reprobado");
                return 0;
            } else {
                System.out.println("Error en las notas ingresadas\n");
            }
        }
        return 0;
    }

    public String estudianteConCinco(double nota, String nombre) {
        /*metodo para devolver el nombre del estudiante que obtuvo 5 en definitiva */
        if (nota == 5) {
            return nombre;
        }
        return null;
    }

    public int resultadoGeneral(int sumAprobados, int numEstudiantes, double sumDef, double desvicionEstandar, Sheet sheet1,int rowNun) {
        int rowNum=rowNun;
        Row row = sheet1.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue("Resultados Generales");
        System.out.println("------------------Resultados Generales----------------------------------------------");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("La cantidad de estudiantes que aprobaron:");
        System.out.println("La cantidad de estudiantes que aprobaron:");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue(sumAprobados);
        System.out.println(sumAprobados);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("La cantidad de estudiantes que no aprobaron:");
        System.out.println("La cantidad de estudiantes que no aprobaron:");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue(numEstudiantes - sumAprobados);
        System.out.println(numEstudiantes - sumAprobados);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("El promedio general de las notas definitivas:");
        System.out.println("El promedio general de las notas definitivas:");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue(sumDef / numEstudiantes);
        System.out.println(sumDef / numEstudiantes);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Desvicion estandar: ");
        System.out.println("Desvicion estandar: ");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue(desvicionEstandar);
        System.out.println(desvicionEstandar + "\n");
        return rowNum;
    }

    public int notasMaximasMinimas(double minimaQuiz, double minimaTaller, double minimaParcial, double maximaTaller, double maximaQuiz, double maximaParcial,Sheet sheet1,int rowNuw) {
        int rowNum=rowNuw;
        Row row = sheet1.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue("Notas Maximas y minimas");

        System.out.println("-----------------------Notas Maximas y minimas----------------------------------------");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Notas mas altas:");
        System.out.println("Notas mas altas:");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Parcial: " + maximaParcial);
        System.out.println("Parcial: " + maximaParcial);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("quiz: " + maximaQuiz);
        System.out.println("quiz: " + maximaQuiz);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("taller: " + maximaTaller );
        System.out.println("taller: " + maximaTaller + "\n");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Notas mas bajas:");
        System.out.println("Notas mas bajas:");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Parcial: " + minimaParcial);
        System.out.println("Parcial: " + minimaParcial);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("quiz: " + minimaQuiz);
        System.out.println("quiz: " + minimaQuiz);
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("taller: " + minimaTaller);
        System.out.println("taller: " + minimaTaller+"\n\n");
        return rowNum;
    }

    public static boolean isValidUsername(String name)
    {
        String regex = "^[A-Za-z]\\w{1,20}$";
        Pattern p = Pattern.compile(regex);
        if (name == null) {
            return false;
        }
        Matcher m = p.matcher(name);
        return m.matches();
    }

    public Sheet principal(int numeroEstudiantes, Estudiante arr[],Sheet sheet1){
        long inicio = System.nanoTime();
        long noCountTime=0;
        int numEstudiantes = numeroEstudiantes;
        if (numEstudiantes==0) {
            do {
                System.out.println("Numero de estudiantes (maximo 1000 estudiantes):");
                long stop1=System.nanoTime();
                String numEstudiantesStr = entrada.next();
                long restart1 =System.nanoTime();
                noCountTime+=restart1-stop1;
                numEstudiantes = isNumericInt(numEstudiantesStr);
                if (numEstudiantes <= 0 || numEstudiantes > 1000) {
                    System.out.println("Error!, introducir numero entre 1 y 1000");
                }
            } while (numEstudiantes <= 0 || numEstudiantes > 1000);

        } else {
              resultados(numeroEstudiantes,arr,true,sheet1);
        }
        return sheet1;
    }

    public Sheet resultados(int numeroEstudiantes,Estudiante arr[], boolean excel,Sheet sheet1){
        long inicio = System.nanoTime(),noCountTime=0;
        int numEstudiantes = numeroEstudiantes, numAprobados = 0, sumAprobados = 0, contador = 0;
        double minimaQuiz = 5, minimaTaller = 5, minimaParcial = 5, maximaTaller = 0, maximaQuiz = 0, maximaParcial = 0, sumDef = 0, varianza = 0, promedio = 0;
        String[] cinco = new String[numEstudiantes];
        String estudiante;

        if (!excel){
            for(int i=0;i< numEstudiantes;i++) {
                long stoptime = System.nanoTime();
                IngreseNotas(i + 1);
                long restart = System.nanoTime();
                noCountTime += restart - stoptime;
                arr[i] = new Estudiante(nombre, nota1,quiz1,quiz2,quiz3,taller1,taller2);
            }
        }
        for (int i = 0; i < arr.length; i++) {
            sumDef = sumDef + arr[i].getNota();
            numAprobados = Mensaje(i,arr);
            sumAprobados = sumAprobados + numAprobados;
            estudiante = estudianteConCinco(arr[i].getNota(),arr[i].getNombre());
            minimaTaller = Math.min(arr[i].getNotaTaller(), minimaTaller);
            minimaParcial = Math.min(arr[i].getNotaParcial(), minimaParcial);
            minimaQuiz = Math.min(arr[i].getNotaQuiz(), minimaQuiz);
            maximaTaller = Math.max(arr[i].getNotaTaller(), maximaTaller);
            maximaQuiz = Math.max(arr[i].getNotaQuiz(), maximaQuiz);
            maximaParcial = Math.max(arr[i].getNotaParcial(), maximaParcial);
            if (estudiante != null) {
                cinco[contador] = estudiante;
                contador += 1;
            }
        }
        promedio = sumDef / numEstudiantes;
        for (int i = 0; i < numEstudiantes; i++) {
            double nota = arr[i].getNota();
            double num = nota - promedio;
            varianza = varianza + (num * num);
        }
        double desviacionEstandar = Math.sqrt(varianza / numEstudiantes);
        int rowNum=0;
        Row row;
        Cell cell;

        System.out.println("------------------------Lista estudiante con cinco--------------------------");
        row = sheet1.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Lista de estudiantes con cinco en definitiva:");
        System.out.println("Lista de estudiantes con cinco en definitiva: ");
        for (int i = 0; i < cinco.length; i++) {
            if (cinco[i] != null) {
                row = sheet1.createRow(rowNum++);
                cell = row.createCell(0);
                cell.setCellValue(i + 1 + ". " + cinco[i]);
                System.out.println(i + 1 + ". " + cinco[i]);
            }
        }

        System.out.println("\n");
        rowNum=notasMaximasMinimas(minimaQuiz, minimaTaller, minimaParcial, maximaTaller, maximaQuiz, maximaParcial,sheet1,rowNum);
        rowNum=resultadoGeneral(sumAprobados, numEstudiantes, sumDef, desviacionEstandar,sheet1,rowNum);
        rowNum=Ordenar(arr, numEstudiantes,sheet1,rowNum);
        long fin = System.nanoTime();
        System.out.println("Duracion: " + ((fin-inicio)-noCountTime)/1e6 + " ms");
        return sheet1;
    }

    static int Ordenar(Estudiante arr[], int n,Sheet sheet1,int rowNun) {
        int rowNum=rowNun;
        int i, j;
        double temp;
        String name;
        boolean swapped;
        for (i = 0; i < n - 1; i++)
        {
            swapped = false;
            for (j = 0; j < n - i - 1; j++)
            {
                if (arr[j].getNota() > arr[j + 1].getNota())
                {
                    // swap arr[j] and arr[j+1]
                    temp = arr[j].getNota();
                    name = arr[j].getNombre();
                    arr[j].setNota(arr[j + 1].getNota());
                    arr[j].setNombre(arr[j + 1].getNombre());
                    arr[j + 1].setNota(temp);
                    arr[j + 1].setNombre(name);
                    swapped = true;
                }
            }
            // IF no two elements were
            // swapped by inner loop, then break
            if (!swapped)
                break;
        }
        System.out.println("---------------------Lista ordenada----------------------------------");
        Row row = sheet1.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue("Lista ordenada por promedio de mayor a menor :");
        System.out.println("Lista ordenada por promedio de mayor a menor :");
        int count=0;
        for (int h = arr.length; h > 0; h--) {
            count+=1;
            row = sheet1.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue(count+". Nombre: " + arr[h-1].getNombre() + "   Nota: " + arr[h-1].getNota());
            System.out.println(count+". Nombre: " + arr[h-1].getNombre() + "   Nota: " + arr[h-1].getNota());
        }
        System.out.println("\n");
        return rowNum;
    }
}
class Estudiante{
    private String nombre,estado;
    private double nota,notaQuiz,quiz1,quiz2,quiz3,notaTaller,taller1,taller2,notaParcial;
    public Estudiante(String nom, double notaParcial,double quiz1,double quiz2,double quiz3,double taller1,double taller2){
        this.nombre = nom;
        this.notaParcial = notaParcial;
        this.quiz1=quiz1;
        this.quiz2=quiz2;
        this.quiz3=quiz3;
        this.taller1=taller1;
        this.taller2=taller2;
        this.notaQuiz= (quiz1+quiz2+quiz3)/3;
        this.notaTaller =(taller1+taller2)/2;
        this.nota= (notaQuiz*0.3)+(notaParcial*0.5)+(notaTaller*0.2);
        if (nota>=3 && nota<=5){
            this.estado= "aprobado";
        } else if (nota>=0 && nota<3) {
            this.estado= "reprobado";
        }
    }
    public double getNota(){
        return nota;
    }
    public String getNombre(){return nombre;}
    public double getNotaQuiz(){return notaQuiz;}
    public double getNotaTaller(){return notaTaller;}
    public double getNotaParcial(){return notaParcial;}

    public String getEstado(){return  estado;}
    public void setNombre(String name){
        this.nombre= name;
    }
    public  void setNota(double note){
        this.nota= note;
    }
    public void setQuiz1(double quiz1){this.quiz1=quiz1;}
    public void setQuiz2(double quiz2){this.quiz2=quiz2;}
    public void setQuiz3(double quiz3){this.quiz3=quiz3;}
    public void setTaller1(double taller1){this.taller1=taller1;}
    public void setTaller2(double taller2){this.taller2=taller2;}
}
