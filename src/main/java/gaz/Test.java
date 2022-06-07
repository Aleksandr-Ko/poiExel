package gaz;

public class Test {
    public static void main(String[] args) {
        Test t = new Test();
        double d = 2111.05;
        RussianMoney_2 r = new RussianMoney_2(d);

        System.out.println(d + " " +r.num2str());
    }
}
