package xlreader;

public class ReadExcel {
    private String Name;
    private int Age;
    private int Marks;

    public ReadExcel() {
    }

    public String getName() {
        return Name;
    }

    public void setName(String name) {
        Name = name;
    }

    public int getAge() {
        return Age;
    }

    public void setAge(int age) {
        Age = age;
    }

    public int getMarks() {
        return Marks;
    }

    public void setMarks(int marks) {
        Marks = marks;
    }

    @Override
    public String toString() {
        return "ReadExcel{" +
                "Name='" + Name + '\'' +
                ", Age=" + Age +
                ", Marks=" + Marks +
                '}';
    }
}
