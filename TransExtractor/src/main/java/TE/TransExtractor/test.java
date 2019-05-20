package TE.TransExtractor;

import java.util.ArrayList;
import java.util.List;

public class test {

	public static void main(String[] args) {


		List<TradosBeam> tbList = new ArrayList<TradosBeam>();
		
		TradosBeam a = new TradosBeam();
		
		a.setChecked("d");
		
		tbList.add(a);
		
		for (TradosBeam tradosBeam : tbList) {
			System.out.println(tradosBeam.getChecked());
		}
		
		

	}

}
