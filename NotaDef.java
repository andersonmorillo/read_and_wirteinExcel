import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

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
/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class NotaDef {

    double quiz1,quiz2,quiz3,taller1,taller2;
    double nota1, nota2, nota3;
    double acu1, acu2, acu3, def;
    String nombre;
    Scanner entrada = new Scanner(System.in);

    Estudiante arr[];

    String reportes[];


    public static final String SAMPLE_XLS_FILE_PATH = "./sample-xls-file.xls";
    public static final String SAMPLE_XLSX_FILE_PATH = "D:\\Anderson\\Downloads\\DATOS.xlsx";
    public static void main(String[] args) throws IOException, InvalidFormatException {
            NotaDef fc = new NotaDef();
            Estudiante arr[] = new Estudiante[0];
            String menu;
            Map <Integer, Estudiante> data= new HashMap<>();

            /*
            System.out.println("-------------------------------------MENU----------------------------------------- ");
            System.out.println("Opciones:");
            System.out.println("1. Introducir N numero de estudiantes (Digitar 1)");
            System.out.println("2. Probar software con datos 10 Estudiantes (Digitar 2)");
            System.out.println("3. Introducir cualquier digito diferente de 1 y 2 para (salir)");
            menu= fc.entrada.next();
            switch (menu){
                case "1":
                    fc.principal(0,arr);
                    break;
                case "2":
                    fc.principal(10,arr);
                    break;
                default:
            }
*/

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }
        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */
        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        /*
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                printCellValue(cell);
                //String cellValue = dataFormatter.formatCellValue(cell);
                //System.out.print(cellValue + "\t");
            }
            System.out.println();
        }*/
        int numEstudiantes=sheet.getLastRowNum();
        arr= new Estudiante[numEstudiantes];
        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        int countRow=0;
        for (Row row: sheet) {
            int countCell=0;
            //arr[countRow]=new Estudiante()
            //data.put(countRow, Estudiante);
            for(Cell cell: row) {
                cell = getCellValue(cell,data,countRow,countCell);
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
            try{
                arr[countRow] = new Estudiante(fc.nombre,fc.nota1, fc.quiz1, fc.quiz2, fc.quiz3, fc.taller1,fc.taller2);
                //data.put(countRow,new Estudiante(fc.nombre,fc.nota1, fc.quiz1, fc.quiz2, fc.quiz3, fc.taller1,fc.taller2));
                countRow++;
            } catch (Exception e){
                break;
            }
            System.out.println(" ");
        }
        // Closing the workbook
        fc.principal(arr.length,arr);
        //workbook.close();


    }


        private static Cell getCellValue(Cell cell, Map <Integer, Estudiante> data,int conRow, int conCell) {
        /*

        * */
        final double cero =0;
        double numero=0;
        switch (cell.getCellType()) {
            case BOOLEAN:
                cell.setCellValue(cero);
                System.out.print(cell.getBooleanCellValue());

                return cell;
            case STRING:
                double res= numberValidation(cell.getStringCellValue());
                if (conCell>=2) {
                    cell.setCellValue(cero);//crear reporte
                } else {
                    if (res!=-1){
                        cell.setCellValue(res);

                        System.out.println(cell.getNumericCellValue());
                    } else {
                        String word = cell.getRichStringCellValue().getString();
                        if (word.length() < 4) {//arreglar valores tipo string en las notas
                            cell.setCellValue(cero);
                            System.out.print(cell.getNumericCellValue());
                        } else {
                            System.out.print(cell.getStringCellValue());


                        }
                    }
                }

                return cell;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.setCellValue(cero);
                } else {
                    numero=cell.getNumericCellValue();
                    if (numero>5 || numero<0) {
                        cell.setCellValue(cero);
                    }else {
                        cell.setCellValue(numero);
                    }
                }
                System.out.println(cell.getNumericCellValue());
                return cell;
            case FORMULA:
                cell.setCellValue(cero);
                System.out.print(cell.getCellFormula());
                return cell;
            case BLANK:
                cell.setCellValue(cero);
                System.out.println(cell.getNumericCellValue());
                //recordar hacer los reportes para los casos especiales
                return  cell;
            default:
                System.out.print("print default\n ");
        }
        System.out.print("\t");
        return null;
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
    public void resultadoGeneral(int sumAprobados, int numEstudiantes, double sumDef, double desvicionEstandar) {
        System.out.println("------------------Resultados Generales----------------------------------------------");
        System.out.println("La cantidad de estudiantes que aprobaron:");
        System.out.println(sumAprobados);
        System.out.println("La cantidad de estudiantes que no aprobaron:");
        System.out.println(numEstudiantes - sumAprobados);
        System.out.println("El promedio general de las notas definitivas:");
        System.out.println(sumDef / numEstudiantes);
        System.out.println("Desvicion estandar: ");
        System.out.println(desvicionEstandar + "\n");
    }
    public void notasMaximasMinimas(double minimaQuiz, double minimaTaller, double minimaParcial, double maximaTaller, double maximaQuiz, double maximaParcial) {
        System.out.println("-----------------------Notas Maximas y minimas----------------------------------------");
        System.out.println("Notas mas altas:");
        System.out.println("Parcial: " + maximaParcial);
        System.out.println("quiz: " + maximaQuiz);
        System.out.println("taller: " + maximaTaller + "\n");
        System.out.println("Notas mas bajas:");
        System.out.println("Parcial: " + minimaParcial);
        System.out.println("quiz: " + minimaQuiz);
        System.out.println("taller: " + minimaTaller+"\n\n");
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
    public void principal(int numeroEstudiantes, Estudiante arr[]){

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
            //for (int i=0;i< arr.length;i++){
               // System.out.println("row: "+arr[i].getNombre()+"     Nota definitiva: "+arr[i].getNota()+ "          quiz1 "+arr[i].getNotaQuiz()+"  Parcial: "+arr[i].getNotaParcial()+ "   Taller: "+ arr[i].getNotaTaller());
            //}
              resultados(numeroEstudiantes,arr,true);
        }
    }


    public void  resultados(int numeroEstudiantes,Estudiante arr[], boolean excel){
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
        resultadoGeneral(sumAprobados, numEstudiantes, sumDef, desviacionEstandar);
        System.out.println("------------------------Lista estudiante con cinco--------------------------");
        System.out.println("Lista de estudiantes con cinco en definitiva: ");
        for (int i = 0; i < cinco.length; i++) {
            if (cinco[i] != null) {
                System.out.println(i + 1 + ". " + cinco[i]);
            }
        }
        System.out.println("\n");
        notasMaximasMinimas(minimaQuiz, minimaTaller, minimaParcial, maximaTaller, maximaQuiz, maximaParcial);
        Ordenar(arr, numEstudiantes);

        long fin = System.nanoTime();
        System.out.println("Duracion: " + ((fin-inicio)-noCountTime)/1e6 + " ms");
    }

    static void Ordenar(Estudiante arr[], int n) {
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
        System.out.println("Lista ordenada por promedio de mayor a menor :");
        int count=0;
        for (int h = arr.length; h > 0; h--) {
            count+=1;
            System.out.println(count+". Nombre: " + arr[h-1].getNombre() + "   Nota: " + arr[h-1].getNota());

        }
        System.out.println("\n");
    }
}
class Estudiante{
    private String nombre;
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

    }
    public double getNota(){
        return nota;
    }
    public String getNombre(){return nombre;}
    public double getNotaQuiz(){return notaQuiz;}
    public double getNotaTaller(){return notaTaller;}
    public double getNotaParcial(){return notaParcial;}
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
