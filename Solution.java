/**
 * The program 'Solution' is used to find primes in an excell file, where the values
 * are stored in first sheet, column B. Program extracts numbers from a file, which
 * is passed to it as program parameter, and prints only primes within those numbers.
 *
 * @author Jakub Karpíšek
 */

import java.io.IOException;
import java.util.ArrayList;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Solution {
    private static final int SHEET_NBR = 0;     // first sheet
    private static final int COLUMN_NBR = 1;    // column B
    private static final int CONSTRAINT_LOWEST_INT_ACCEPTED = 1;

    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("Usage: java -jar Solution.jar fileName");
            return;
        }
        try {
            findSolutionInFile(args[0]);
        } catch (IOException e) {
            System.out.printf("File '%s' not found.%n", args[0]);
        }
    }

    /**
     * Tries to open file, parse data and print results.
     * @param fileName
     * @throws IOException if file does not exist
     */
    private static void findSolutionInFile(String fileName) throws IOException {
        File file = new File(fileName);   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        ArrayList<Integer> extractedNumbers = extractNumbers(wb.getSheetAt(SHEET_NBR));
        ArrayList<Integer> primes = getPrimes(extractedNumbers);
        displayResult(primes);
    }

    /**
     * Print one number per line to stdout.
     * @param primes
     */
    private static void displayResult(ArrayList<Integer> primes) {
        for (Integer i : primes) {
            System.out.println(i);
        }
    }

    /**
     * Finds prime numbers within list and returns them i another list.
     * @param allNumbers input list
     * @return list of primes
     */
    private static ArrayList<Integer> getPrimes(ArrayList<Integer> allNumbers) {
        ArrayList<Integer> primes = new ArrayList<>();
        for (Integer number : allNumbers) {
            if (isPrime(number)) {
                primes.add(number);
            }
        }
        return primes;
    }

    /**
     * Check if given number is prime
     * tests only numbers from 2 to square value of input number
     * @param number
     * @return true if number is prime
     */
    private static boolean isPrime(int number) {
        int maxToTry = (int) Math.sqrt(number);    // test only numbers from 2 to sqrt(number)
        for (int i = 2; i <= maxToTry; i++) {
            if (number%i == 0) {     //not a prime number
                return false;
            }
        }
        return true;
    }

    /**
     * Reads numbers from a column in excel sheet and return them in a list.
     * @param sheet input excel sheet
     * @return list of numbers
     */
    private static ArrayList<Integer> extractNumbers(XSSFSheet sheet) {
        Iterator<Row> itr = sheet.iterator();    //iterating over sheet rows
        ArrayList<Integer> numbersToReturn = new ArrayList<>();
        while (itr.hasNext())
        {
            String cellValue = getCellValue(itr.next(), COLUMN_NBR);
            addValueToListIfInteger(cellValue, numbersToReturn);
        }
        return numbersToReturn;
    }

    /**
     * Validate cell value and add to list of numbers if valid.
     * @param cellValue value to validate
     * @param numbersToReturn list to append the value to
     */
    private static void addValueToListIfInteger(String cellValue, ArrayList<Integer> numbersToReturn) {
        try {
            validateAndAdd(cellValue, numbersToReturn);
        } catch (NumberFormatException e) {
            // cell is not an integer
        }
    }

    /**
     * Add Integer value of cell to list if it is positive integer
     * @param cellValue
     * @param numbersToReturn
     * @throws NumberFormatException when cell does not contain integer
     */
    private static void validateAndAdd(String cellValue, ArrayList<Integer> numbersToReturn) throws NumberFormatException {
        int nbr = Integer.parseInt(cellValue);
        if (nbr >= CONSTRAINT_LOWEST_INT_ACCEPTED) {
            numbersToReturn.add(nbr);
        } else {
//            System.out.printf("Negative number: '%d'%n", nbr);
        }
    }

    /**
     * Read cell value as String
     * @param row
     * @param col
     * @return String value of cell
     */
    private static String getCellValue(Row row, int col) {
        Cell cell = row.getCell(col);
        return cell.getStringCellValue();
    }
}