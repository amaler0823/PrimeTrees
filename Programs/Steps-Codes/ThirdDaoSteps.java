package com;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.TreeMap;

public class ThirdDaoSteps {
    int row = 900; //for {1,2,3...404550}
    TreeMap<Integer, Integer> thrDaoMap1 = new TreeMap<>();
    TreeMap<Integer, String> thrDaoMap2 = new TreeMap<>();
    StringBuffer steps = new StringBuffer();

    public ThirdDaoSteps() {
    }

    public static void main(String[] args) {
        ThirdDaoSteps thirdDaoSteps = new ThirdDaoSteps();
        thirdDaoSteps.dataList();
        thirdDaoSteps.printSteps();
    }

    void printSteps() {
         /*
         AIS2,...AIS10
         AIS6 to ***
         */

        Iterator<Integer> iterator = thrDaoMap1.keySet().iterator();
        while (iterator.hasNext()) {
            int key1 = iterator.next();
            if (key1 < thrDaoMap1.size() - 1) {
                int val1 = thrDaoMap1.get(key1);
                int val2 = thrDaoMap1.get(key1 + 1);

                String[] strings1 = thrDaoMap2.get(val1).split(",");
                String[] strings2 = thrDaoMap2.get(val2).split(",");
                int diffX = Integer.valueOf(strings2[0]) - Integer.valueOf(strings1[0]);
                int diffY = Integer.valueOf(strings2[1]) - Integer.valueOf(strings1[1]);

                switch (diffX) {
                    case 0://Cannon-Horizontal-Right
                        //AIS2 to val2
                        switch (diffY) {
                            case 2:
//                                steps.append("AIS2 to " + val2 + "@@@");
                                steps.append("2" + "@@@");
                                continue;
                            default://Pause/Stop
//                                steps.append("AIS10 -diffX:" + diffX + "-diffY:" + diffY + " to " + val2 + "@@@");
                                steps.append("10" + "@@@");
                                continue;
                        }
                    case 1://Mandarin3&7//Knight-Horizontal4&8//Soldier-Vertical6
                        switch (diffY) {
                            case 0://Soldier-Vertical-Down
//                                steps.append("AIS6 to " + val2 + "@@@");
                                steps.append("6" + "@@@");
                                continue;
                            case 1://Mandarin-Lower-Right
//                                steps.append("AIS3 to " + val2 + "@@@");
                                steps.append("3" + "@@@");
                                continue;
                            case -1://Mandarin-Lower-Left
//                                steps.append("AIS7 to " + val2 + "@@@");
                                steps.append("7" + "@@@");
                                continue;
                            case 2://Knight-Horizontal-Lower-Right
//                                steps.append("AIS4 to " + val2 + "@@@");
                                steps.append("4" + "@@@");
                                continue;
                            case -2://Knight-Horizontal-Lower-Left
//                                steps.append("AIS8 to " + val2 + "@@@");
                                steps.append("8" + "@@@");
                                continue;
                            default://Pause/Stop
//                                steps.append("AIS10 -diffX:" + diffX + "-diffY:" + diffY + " to " + val2 + "@@@");
                                steps.append("10" + "@@@");
                                continue;
                        }
                    case 2://Knight-Vertical5&9
                        switch (diffY) {
                            case 1://Knight-Vertical-Lower-Right
//                                steps.append("AIS5 to " + val2 + "@@@");
                                steps.append("5" + "@@@");
                                continue;
                            case -1://Knight-Vertical-Lower-Left
//                                steps.append("AIS9 to " + val2 + "@@@");
                                steps.append("9" + "@@@");
                                continue;
                            default://Pause/Stop
//                                steps.append("AIS10 -diffX:" + diffX + "-diffY:" + diffY + " to " + val2 + "@@@");
                                steps.append("10" + "@@@");
                                continue;
                        }
                    default://Pause/Stop
//                        steps.append("AIS10 -diffX:" + diffX + "-diffY:" + diffY + " to " + val2 + "@@@");
                        steps.append("10" + "@@@");
                }
            }
        }
        System.out.println(steps.toString());//steps printing
    }

    private void dataList() {
        List<List<Object>> list = new ArrayList();
        int ma = 1;
        int a = 0;

        int id = 0;

        for (int i = 0; i < this.row; ++i) {
            //third dao f(x) = (x * (x + 1)) / 2 - 15
            //forth dao f(x) = (x * (x + 1)) / 2 - 21
            int row21 = (i * (i + 1)) / 2 - 21;
            int row15 = (i * (i + 1)) / 2 - 15;
            for (int j = 0; j < i; ++j) {
                a = Integer.valueOf(j + ma) + 1;
                if (this.numIsP(a) && a > row21 && a < row15) {
                    thrDaoMap1.put(id, a);++id;
                    thrDaoMap2.put(a, i + "," + (j + 1));
                }
            }
            ma = a;
        }
//        System.out.println(firstDaoMap1);
//        System.out.println(firstDaoMap2);
    }

    Boolean numIsP(int num) {
        Boolean b = true;
        for (int i = 2; i < num; i++) {
            if (BigInteger.ZERO.equals(BigInteger.valueOf(num).mod(BigInteger.valueOf(i)))) {
                return false;
            }
        }
        return b;
    }
}
