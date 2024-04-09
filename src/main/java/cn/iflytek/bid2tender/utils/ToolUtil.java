package cn.iflytek.bid2tender.utils;

public class ToolUtil {

    /**
     * 数字转汉字
     *
     * @param number
     */
    public static String convertNumber(int number) {
        //数字对应的汉字
        String[] num = {"一", "二", "三", "四", "五", "六", "七", "八", "九"};
        //单位
        String[] unit = {"", "十", "百", "千", "万", "十", "百", "千", "亿", "十", "百", "千", "万亿"};
        //将输入数字转换为字符串
        String result = String.valueOf(number);
        //将该字符串分割为数组存放
        char[] ch = result.toCharArray();
        //结果 字符串
        String str = "";
        int length = ch.length;
        for (int i = 0; i < length; i++) {
            int c = (int) ch[i] - 48;
            if (c != 0) {
                if (number > 9 && number < 20 && i == 0) {
                    str += unit[length - i - 1];
                } else {
                    str += num[c - 1] + unit[length - i - 1];
                }
            }
        }
        return str;
    }

}
