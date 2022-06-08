package gaz;

public class Test {
    public static void main(String[] args) {
        Test t = new Test();
        double d1 = 319319319319.01;
        double d2 = 319.01;
        double d3 = 19.01;
        double d4 = 2.01;

        RussianMoney_2 r1 = new RussianMoney_2(d1);
        RussianMoney_2 r2 = new RussianMoney_2(d2);
        RussianMoney_2 r3 = new RussianMoney_2(d3);
        RussianMoney_2 r4 = new RussianMoney_2(d4);

        System.out.println(d1 + " " +r1.num2str());
        System.out.println(d2 + " " +r2.num2str());
        System.out.println(d3 + " " +r3.num2str());
        System.out.println(d4 + " " +r4.num2str());
    }
}
