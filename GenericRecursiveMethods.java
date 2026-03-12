package genericRecursiveMethods;

/* Chuqiao Yan
 * chuqiaoy
 * 119690833
 * 0301*/

/* I pledge on my honor that I have not given or received any unauthorized
 * assistance on this assignment.*/

public class GenericRecursiveMethods {

    public static <T> int numMatching(GLList<T> list, GLList<T> otherList) {
    	if (list == null || otherList == null) {
    		throw new IllegalArgumentException();
    	}else {
    		return helpNumMatching(list, otherList, 0, 0);
    	}
    }
  
    private static <T> int helpNumMatching(GLList<T> list,
    					GLList<T> otherList, int index, int numMatch) {
    	if (list.get(index) == null || list.get(index) == null) {
    		return numMatch;
    	}else if((list.get(index)).equals(otherList.get(index))){
    		return helpNumMatching(list, otherList, index+1, numMatch+1);
    	}else {
    		return helpNumMatching(list, otherList, index+1, numMatch);
    	}
    }

    public static <T extends Comparable<T>> T firstLarger(GLList<T> list,
                                                          T element) {
    	if(list == null || element == null) {
    		throw new IllegalArgumentException();
    	}else {
    		return helpFirstLarger(list, element, 0);
    	}
    }
    
    private static <T extends Comparable<T>> T helpFirstLarger
    							(GLList<T> list, T element, int index) {
    	if(list.get(index) == null) {
    		return null;
    	}else if(list.get(index).compareTo(element) > 0) {
    		return list.get(index); 
    	}else {
    		return helpFirstLarger(list, element, index + 1);
    	}
    	
    }

    public static <T> GLList<T> insertAfter(GLList<T> list, T element,
                                            T newElt) {
    	if(list == null || element == null || newElt == null) {
    		return null;
    	}else {
    		GLList<T> newList =new GLList<>();
    		return helpInsertAfter(list, element, newElt, newList, 0);
    	}
    }
    
    private static <T> GLList<T> helpInsertAfter(GLList<T> list, T element,
            T newElt,GLList<T> newList, int index) {
    	if (list.get(index) == null) {
    		return newList;
    	}else {
    		newList.append(list.get(index));
    		if (list.get(index).equals(element)) {
    			newList.append(newElt);
    		}
    		return helpInsertAfter
    						(list, element, newElt, newList, index + 1);
    	}
    }
    

    public static <T> GLList<T> invert(GLList<T> list) {
    	if(list == null) {
    		return null;
    	}else {
    		GLList<T> newList =new GLList<>();
    		return helpInvert(list, list.size() - 1, newList);
    	}
    }
    private static <T> GLList<T> 
    			helpInvert(GLList<T> list, int index, GLList<T> newList) {
    	if (index < 0) {
    		return newList;
    	}else {
    		newList.append(list.get(index));
    		return helpInvert(list, index - 1, newList);
    	}
    	
    }

}
