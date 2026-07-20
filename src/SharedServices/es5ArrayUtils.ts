// ES5 compatible utility functions for array operations
// Replaces ES2015+ features like Set for SPFx projects

export class ES5ArrayUtils {
  
  /**
   * Remove duplicate values from array (replaces new Set + Array.from)
   * @param array - Array to deduplicate
   * @returns Array with unique values
   */
  static getUniqueValues<T>(array: T[]): T[] {
    return array.filter((value, index) => {
      return array.indexOf(value) === index;
    });
  }

  /**
   * Remove duplicate objects from array based on a property
   * @param array - Array of objects to deduplicate  
   * @param property - Property name to check for uniqueness
   * @returns Array with unique objects
   */
  static getUniqueObjects<T>(array: T[], property: keyof T): T[] {
    const seen: any[] = [];
    return array.filter(item => {
      const value = item[property];
      if (seen.indexOf(value) === -1) {
        seen.push(value);
        return true;
      }
      return false;
    });
  }

  /**
   * Merge arrays and remove duplicates (replaces [...new Set([...array1, ...array2])])
   * @param arrays - Arrays to merge and deduplicate
   * @returns Merged array with unique values
   */
  static mergeAndDeduplicate<T>(...arrays: T[][]): T[] {
    let result: T[] = [];
    
    // Concatenate all arrays
    arrays.forEach(array => {
      result = result.concat(array);
    });
    
    // Remove duplicates
    return this.getUniqueValues(result);
  }

  /**
   * Get unique non-empty values from array with filtering
   * @param array - Input array
   * @param filterFn - Optional filter function
   * @returns Unique filtered values
   */
  static getUniqueFilteredValues<T>(array: T[], filterFn?: (value: T) => boolean): T[] {
    let filtered = array;
    
    if (filterFn) {
      filtered = array.filter(filterFn);
    }
    
    return this.getUniqueValues(filtered);
  }
}