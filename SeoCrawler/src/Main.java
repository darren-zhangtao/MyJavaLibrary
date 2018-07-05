
public class Main {

    public static void main(String[] args) {
        String []urlData;
        urlData = Util.readExcel();
        int urlDataLength = urlData.length;
        String [][]crawlerData = new String[urlDataLength][];
        for (int i=0; i< urlDataLength; i++) {
            if (urlData[i] != null) {
                crawlerData[i] = Util.GetUrlData(urlData[i]);
            }
        }
        Util.writeExcel(crawlerData);
    }
}