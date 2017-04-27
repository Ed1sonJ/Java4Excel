package org.dxj.excel.bean;

/**
 * @author : duxiji
 * @date : 2017/4/27
 * @description : 装载学生的排名
 */
public class Rank {
    private String id;//学号
    private String chinese;
    private String math;
    private String english;
    private String physics;
    private String chemology;
    private String biological;
    private String politics;
    private String history;
    private String geography;
    private String wenkeRank;
    private String likeRank;
    private String totalRank;//9科总排名

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getChinese() {
        return chinese;
    }

    public void setChinese(String chinese) {
        this.chinese = chinese;
    }

    public String getMath() {
        return math;
    }

    public void setMath(String math) {
        this.math = math;
    }

    public String getEnglish() {
        return english;
    }

    public void setEnglish(String english) {
        this.english = english;
    }

    public String getPhysics() {
        return physics;
    }

    public void setPhysics(String physics) {
        this.physics = physics;
    }

    public String getChemology() {
        return chemology;
    }

    public void setChemology(String chemology) {
        this.chemology = chemology;
    }

    public String getBiological() {
        return biological;
    }

    public void setBiological(String biological) {
        this.biological = biological;
    }

    public String getPolitics() {
        return politics;
    }

    public void setPolitics(String politics) {
        this.politics = politics;
    }

    public String getHistory() {
        return history;
    }

    public void setHistory(String history) {
        this.history = history;
    }

    public String getGeography() {
        return geography;
    }

    public void setGeography(String geography) {
        this.geography = geography;
    }

    public String getWenkeRank() {
        return wenkeRank;
    }

    public void setWenkeRank(String wenkeRank) {
        this.wenkeRank = wenkeRank;
    }

    public String getLikeRank() {
        return likeRank;
    }

    public void setLikeRank(String likeRank) {
        this.likeRank = likeRank;
    }

    public String getTotalRank() {
        return totalRank;
    }

    public void setTotalRank(String totalRank) {
        this.totalRank = totalRank;
    }

    /**
     * 获取9科的平均排名
     *
     * @param rank
     * @return
     */
    public static double get9rank(Rank rank) {
        double rankTotal = (Float.parseFloat(rank.getChinese()) + Float.parseFloat(rank.getMath()) + Float.parseFloat(rank.getEnglish()) + Float.parseFloat(rank.getPhysics())
                + Float.parseFloat(rank.getChemology()) + Float.parseFloat(rank.getBiological()) + Float.parseFloat(rank.getPolitics())
                + Float.parseFloat(rank.getHistory()) + Float.parseFloat(rank.getGeography())) / 9;
        rankTotal = Double.parseDouble(String.format("%.2f", rankTotal));
        return rankTotal;
    }
}
