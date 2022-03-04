public class ParcedExel {
    private String num;
    private String sfera;
    private String osnovanie;
    private String vc;
    private String startDate;
    private String endDate;
    private String parsedDate;
    private String ispolnitel;

    public ParcedExel(String num, String sfera, String osnovanie, String vc, String startDate, String endDate, String ispolnitel) {
        this.num = num;
        this.sfera = sfera;
        this.osnovanie = osnovanie;
        this.vc = vc;
        this.startDate = startDate;
        this.endDate = endDate;
        String[] splits = startDate.split("\\.");
        this.parsedDate = "«" + splits[0] + "»" + " " + Maps.getMonths().get(Integer.parseInt(splits[1]) - 1) + " " + splits[2];
        this.ispolnitel = Maps.getIspolniteli().get(ispolnitel);
    }

    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public String getSfera() {
        return sfera;
    }

    public void setSfera(String sfera) {
        this.sfera = sfera;
    }

    public String getOsnovanie() {
        return osnovanie;
    }

    public void setOsnovanie(String osnovanie) {
        this.osnovanie = osnovanie;
    }

    public String getVc() {
        return vc;
    }

    public void setVc(String vc) {
        this.vc = vc;
    }

    public String getStartDate() {
        return startDate;
    }

    public void setStartDate(String startDate) {
        this.startDate = startDate;
    }

    public String getEndDate() {
        return endDate;
    }

    public void setEndDate(String endDate) {
        this.endDate = endDate;
    }

    public String getParsedDate() {
        return parsedDate;
    }

    public void setParsedDate(String parsedDate) {
        this.parsedDate = parsedDate;
    }

    public String getIspolnitel() {
        return ispolnitel;
    }

    public void setIspolnitel(String ispolnitel) {
        this.ispolnitel = ispolnitel;
    }
}
