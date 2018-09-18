package OfficeToolTest;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 9/18/2018 10:06 PM
 */
public class TestBean {
    private String test;


    /**
     * Gets test.
     *
     * @return Value of test.
     */
    public String getTest() {
        return test;
    }

    /**
     * Sets new test.
     *
     * @param test New value of test.
     */
    public void setTest(String test) {
        this.test = test;
    }


    @Override
    public String toString() {
        return "TestBean{" +
                "test='" + test + '\'' +
                '}';
    }
}
