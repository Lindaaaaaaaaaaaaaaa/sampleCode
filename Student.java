package spss;
//Chuqiao Yan,119690833, chuqiaoy,TA: Quinn 0301 

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

/*This class records all informations for a student. Each student have a name,
 * the number of times they have submitted work, their highest total, and
 * their scores.This class mainly records all the information for the fields
 * and also updates the number of times they have submitted work, their
 * highest total, and their scores. The scores are updated only when the
 * new one has a higher total.
*/

public class Student {
	// Fields
	private String name;// store the names
	private int submitted;
	private ArrayList<List> submission;
	private int highestTotal;

	// constructor, initialize the fields.
	public Student(String name) {
		this.name = name;
		submitted = 0;
		highestTotal = 0;
		submission = new ArrayList();

	}

	// return name
	public String getName() {
		return name;
	}

	// return submitted
	public int getSubmitted() {
		return submitted;
	}

	// return higher
	public int getHigher() {
		return highestTotal;
	}

	// increase submit by 1
	public void increaseubmitted() {
		submitted++;
	}

	// check if the total is higher then previous
	public boolean highest(int total) {
		boolean higher = false;
		if (total > highestTotal) {
			highestTotal = total;
			higher = true;
		}
		return higher;

	}

	// store test results
	public void storeResults(List<Integer> list, int total) {
		submission.add(list);
		if (total > highestTotal) {
			highestTotal = total;
		}
	}

	// tests passed
	public int testPassed() {
		int passed = 0;
		if (submission.size() > 0) {

			// obtaining the information
			List<Integer> results = new ArrayList();
			int a = submission.size() - 1;
			results.addAll(submission.get(a));

			// for loop going through the scores of the student,and recording
			// how many times they had passed
			for (int i = 0; i < results.size(); i++) {
				if (results.get(i) > 0) {
					passed++;
				}
			}
		}
		return passed;
	}

//to string method used for testing purposes
	public String toString() {
		String info = "";
		info += "name" + name + "\n";
		info += "submission" + submission.toString() + "\n";
		info += "submitted" + submitted + "\n";
		info += "highestTotal" + highestTotal;

		return info;

	}
}
