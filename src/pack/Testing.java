package pack;

public class Testing {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Testing Testing = new Testing();
		System.out.println("Change done by user 1 on branch 1 :: "+Testing.getClass().getName());
		System.out.println("Changes by user 2 @ branch 1");
		System.out.println("Changes by user 1 @ branch 1");
		System.out.println("Testing remote branch");
		System.out.println("Conflict test by another user");
		System.out.println("Testing for tag release");
	}

}
