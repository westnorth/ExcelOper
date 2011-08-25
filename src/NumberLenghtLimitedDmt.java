import javax.swing.text.*;

public class NumberLenghtLimitedDmt extends PlainDocument {

	/**
	 * 限制JTextField只能输入数字及小数点
	 */
	private static final long serialVersionUID = 1L;
	private int limit;

	public NumberLenghtLimitedDmt(int limit) {
		super();
		this.limit = limit;
	}
	public NumberLenghtLimitedDmt()
	{
		super();
		this.limit = 20;
	}

	public void insertString(int offset, String str, AttributeSet attr)
			throws BadLocationException {
		if (str == null) {
			return;
		}
		if ((getLength() + str.length()) <= limit) {

			char[] upper = str.toCharArray();
			int length = 0;
			for (int i = 0; i < upper.length; i++) {
				//内容只能是数字或小数点
				if (upper[i] >= '0' && upper[i] <= '9') {
					upper[length++] = upper[i];
				}
			}
			super.insertString(offset, new String(upper, 0, length), attr);
		}
	}
}