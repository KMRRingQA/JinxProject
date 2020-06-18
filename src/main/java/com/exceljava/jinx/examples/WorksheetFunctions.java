package com.exceljava.jinx.examples;

import com.exceljava.jinx.*;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Example worksheet functions (UDFs) that can be called
 * from Excel via Jinx.
 *
 * See the examples spreadsheet included in the Jinx download.
 */
public class WorksheetFunctions {
    private final ExcelAddIn xl;

    /**
     * A non-default constructor is not required, but if there is one that
     * takes an ExcelAddIn then it will be used when instantiating the class
     * (which is only done if there are non-static methods).
     *
     * @param xl ExcelAddIn object for calling back into Excel and Jinx
     */
    public WorksheetFunctions(ExcelAddIn xl) {
        this.xl = xl;
    }

    /**
     * Worksheet function that returns the Jinx version.
     *
     * This is a volatile function so will be recalculated each time
     * the workbook is opened or refreshed.
     *
     * @return Jinx version
     */
    @ExcelFunction(
            value = "jinx.version",
            isVolatile = true
    )
    public static String getJinxVersion() {
        Package jinx = Package.getPackage("com.exceljava.jinx");
        return jinx.getImplementationVersion();
    }

    /**
     * Worksheet function that returns the Java runtime version.
     *
     * This is a volatile function so will be recalculated each time
     * the workbook is opened or refreshed.
     *
     * @return Java version
     */
    @ExcelFunction(
            value = "java.version",
            isVolatile = true
    )
    public static String getJavaVersion() {
        return System.getProperty("java.version");
    }

    /**
     * Worksheet function that returns the Jinx config file.
     *
     * This is a volatile function so will be recalculated each time
     * the workbook is opened or refreshed.
     *
     * @return Jinx config file
     */
    @ExcelFunction(
            value = "jinx.config",
            isVolatile = true
    )
    public String getJinxConfigPath() {
        return xl.getConfigPath();
    }

    /**
     * Worksheet function that returns the Jinx config file.
     *
     * This is a volatile function so will be recalculated each time
     * the workbook is opened or refreshed.
     *
     * @return Jinx log file
     */
    @ExcelFunction(
            value = "jinx.log",
            isVolatile = true
    )
    public String getJinxLogPath() {
        return xl.getLogPath();
    }

    /**
     * Return (x * y) + z
     *
     * This is a simple function only intended to show how arguments can
     * be passed to a worksheet function and named so they appear in the
     * function wizard.
     */

    @ExcelFunction(
            value = "jinx.multiplyAndAdd",
            description = "Simple function that returns (x * y) + z"
    )
    @ExcelArguments({
            @ExcelArgument("x"),
            @ExcelArgument("y"),
            @ExcelArgument("z")
    })
    public static double multiplyAndAdd(int x, double y, double z) {
        return (x * y) + z;
    }

    @ExcelFunction(
            value = "jinx.Hello",
            description = "returns \"Hello <name> \""
    )
    @ExcelArguments({
            @ExcelArgument("name")
    })
    public static String hello(String name) {
        return "Hello " + name;
    }


    @ExcelFunction(
            value = "jinx.Modifier",
            description = "returns modifier from Ability score"
    )
    @ExcelArguments({
            @ExcelArgument("score")
    })
    public static Integer modifier(Integer score) {
        score = (score-10)/2;

        return score;
    }

    @ExcelFunction(
            value = "jinx.AverageDie",
            description = "returns the result from an average die roll. Input in format 'ndx' where n is number of dice rolled and x the number of sides"
    )
    @ExcelArguments({
            @ExcelArgument("score")
    })
    public static Double average(String input) {
        Integer rolls = Integer.parseInt(input.split("d")[0]);
        Double diceSide = Double.parseDouble(input.split("d")[1]);
        return rolls * (diceSide/2+0.5);
    }

    @ExcelFunction(
            value = "jinx.DamageOutput",
            description = "input hit dc, target ac, crit range (19/20), per hit damage, per crit damage and number of attacks. Returns expected damage per round."
    )
    @ExcelArguments({
            @ExcelArgument("hitDC"),
            @ExcelArgument("targetAC"),
            @ExcelArgument("critRange"),
            @ExcelArgument("perHitDamage"),
            @ExcelArgument("perCritDamage"),
            @ExcelArgument("attackNumber")

    })
    public static Double average(Integer hitDC, Integer targetAC, Integer critRange, Double perHitDamage, Double perCritDamage, Integer attackNumber) {
        double counter = 0;
        for (int i =2; i<critRange; i++){
            if (hitDC+i>=targetAC){
                counter+=1;
            }
        }
        return attackNumber*(perHitDamage*counter/20) + ((20-critRange+1)*perCritDamage)/20;
    }

    @ExcelFunction(
            value = "jinx.chanceToHit",
            description = "Returns the probability to have (a normal) hit against a target."
    )
    @ExcelArguments({
            @ExcelArgument("targetAC"),
            @ExcelArgument("hitDC"),
            @ExcelArgument("critRange (19/20)"),
            @ExcelArgument("normal/advantage/disadvantage/triple")
    })
    public static Double chanceToHit(Integer ac, Integer dc, Integer crit, String type) {
        int counter = 0;
        for (int i = 2; i < 21; i++) {
            if (i + dc >= ac) {
                counter++;
            }
        }
        double hit = counter * 0.05;
        if (type.toLowerCase().equals("advantage")) {
            hit = 1 - (1 - hit) * (1 - hit);
            hit -= 1 - ((1 - (0.05 * (21 - crit))) * (1 - (0.05 * (21 - crit))));
        } else if (type.toLowerCase().equals("disadvantage")) {
            hit = hit * hit;
            hit -= (0.05 * (21 - crit)) * (0.05 * (21 - crit));
        } else if (type.toLowerCase().equals("triple")) {
            hit = 1 - (1 - hit) * (1 - hit) * (1 - hit);
            hit -= 1 - ((1 - (0.05 * (21 - crit))) * (1 - (0.05 * (21 - crit))) * (1 - (0.05 * (21 - crit))));
        } else {
            hit -= 0.05 * (21 - crit);
        }
        if (hit < 0) {
            hit = 0;
        }
        return hit;
    }

    @ExcelFunction(
            value = "jinx.totalAdvantage",
            description = "Returns the better of two sets of rolls. Use for savage attacker."
    )
    @ExcelArguments({
            @ExcelArgument("number of dice (i.e. 2 for greatsword, 1 for other)"),
            @ExcelArgument("sides of dice (i.e. 6 for greatsword)")
    })
    public static double totalAdvantage(int numberDice, int diceSide) {

        List<Integer> rollList = new ArrayList<Integer>();
        Integer sum = 0;
        for (int i = 0; i < 10000000; i++) {

            for (int j = 0; j < 2; j++) {
                Integer tempSum = 0;
                for (int k = 0; k < numberDice; k++) {
                    tempSum += (int) ((Math.random() * (diceSide)) + 1);
                }
                rollList.add(tempSum);
            }
            rollList.sort(Collections.reverseOrder());
            sum += rollList.get(0);
            rollList.clear();
        }

        return (double) sum / 10000000;
    }
    /**
     * Join a list of strings with a delimiter.
     *
     * This demonstrates using a variable number of arguments in an Excel function.
     */
    @ExcelFunction(
            value = "jinx.stringJoin",
            description = "Join multiple strings with a separator"
    )
    @ExcelArguments({
            @ExcelArgument("sep"),
            @ExcelArgument("string")
    })
    public static String stringJoin(String sep, String ... strings) {
        // Note: Could use String.join in Java 8, but these examples are for Java >= 6
        StringBuilder builder = new StringBuilder();
        for (int i=0; i<strings.length; ++i) {
            if (i >= 1) {
                builder.append(sep);
            }
            builder.append(strings[i]);
        }
        return builder.toString();
    }

    /**
     * Transposes a 2d array of numbers.
     * When the transposed array is returned to Excel, the output range is automatically
     * resized as 'autoResize' is set to true.
     */
    @ExcelFunction(
            value = "jinx.transpose",
            description = "Transposes a 2d matrix of numbers",
            autoResize = true
    )
    @ExcelArguments({
            @ExcelArgument("array")
    })
    public static double[][] transpose(double[][] array) {
        int m = array.length;
        int n = array[0].length;

        double[][] transposed = new double[n][m];

        for(int x = 0; x < n; x++) {
            for(int y = 0; y < m; y++) {
                transposed[x][y] = array[y][x];
            }
        }

        return transposed;
    }

    /**
     * As well as simple types, Jinx can handle returning complex
     * objects to the Excel to pass to other functions.
     *
     * Objects that can't be converted to Excel types are cached and a
     * handle to the object is returned to Excel. The cache is managed
     * and when no longer needed the objects are removed from the cache
     * to be garbage collected.
     */
    public class CachedObjectExample {
        private final String name;

        public CachedObjectExample(String name) {
            this.name = name;
        }

        public String getName() {
            return this.name;
        }
    }

    /**
     * Create a new CachedObjectExample instance to demonstrate returning
     * an object to Excel.
     */
    @ExcelFunction("jinx.createCachedObject")
    public CachedObjectExample createCachedObject(String name) {
        return new CachedObjectExample(name);
    }

    /**
     * Takes a CachedObjectExample instance to demonstrate passing an
     * object returned to the Excel sheet.
     */
    @ExcelFunction("jinx.cachedObjectName")
    public String getCachedObjectName(CachedObjectExample obj) {
        return obj.getName();
    }

    /**
     * Gets the address of the calling cell and returns it.
     */
    @ExcelFunction(
            value = "jinx.getCallerAddress",
            isMacroType = true
    )
    public String getCallerAddress() {
        ExcelReference caller = xl.getCaller();
        return caller.getAddress();
    }

    /**
     * Gets the formula of the cell passed into the function.
     */
    @ExcelFunction(
            value = "jinx.getFormula",
            isMacroType = true
    )
    public String getFormula(ExcelReference cell) {
        return cell.getFormula();
    }
}
