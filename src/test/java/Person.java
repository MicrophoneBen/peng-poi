import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * Created by Administrator on 2018-10-20.
 */
public class Person {

    private String name;

    private Integer age;

    private LocalDate birthday;

    private Boolean active;

    private LocalDateTime createDate;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public LocalDate getBirthday() {
        return birthday;
    }

    public void setBirthday(LocalDate birthday) {
        this.birthday = birthday;
    }

    public Boolean getActive() {
        return active;
    }

    public void setActive(Boolean active) {
        this.active = active;
    }

    public LocalDateTime getCreateDate() {
        return createDate;
    }

    public void setCreateDate(LocalDateTime createDate) {
        this.createDate = createDate;
    }

    public Person() {
    }

    public Person(String name, Integer age, LocalDate birthday, Boolean active, LocalDateTime createDate) {
        this.name = name;
        this.age = age;
        this.birthday = birthday;
        this.active = active;
        this.createDate = createDate;
    }
}
