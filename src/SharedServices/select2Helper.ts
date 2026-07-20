import 'select2';
import 'select2/dist/css/select2.min.css';
import * as $ from 'jquery';

export interface Select2Option {
  id: string;
  text: string;
  selected?: boolean;
}

export interface Select2Config {
  placeholder?: string;
  allowClear?: boolean;
  width?: string;
  theme?: string;
  searchable?: boolean;
  minimumInputLength?: number;
  maximumSelectionLength?: number;
  closeOnSelect?: boolean;
  tags?: boolean;
}

export class Select2Helper {
  
  /**
   * Initialize Select2 on a dropdown element
   * @param selector - jQuery selector for the dropdown element
   * @param options - Array of Select2Option objects
   * @param config - Select2 configuration options
   */
  public static initializeSelect2(
    selector: string, 
    options: Select2Option[] = [], 
    config: Select2Config = {}
  ): void {
    const defaultConfig: Select2Config = {
      placeholder: 'Select an option...',
      allowClear: true,
      theme: 'default',
      searchable: true,
      minimumInputLength: 0,
      closeOnSelect: true
    };

    const finalConfig = { ...defaultConfig, ...config };
    
    // Destroy existing Select2 if it exists
    if ($(selector).hasClass('select2-hidden-accessible')) {
      ($(selector) as any).select2('destroy');
    }

    // Clear existing options
    $(selector).empty();
    
    // Add placeholder option if allowClear is true
    if (finalConfig.allowClear) {
      $(selector).append('<option></option>');
    }

    // Add options to the select element
    options.forEach(option => {
      const $option = $(`<option value="${option.id}">${option.text}</option>`);
      if (option.selected) {
        $option.prop('selected', true);
      }
      $(selector).append($option);
    });

    // Initialize Select2
    ($(selector) as any).select2({
      placeholder: finalConfig.placeholder,
      allowClear: finalConfig.allowClear,
      width: finalConfig.width,
      theme: finalConfig.theme,
      minimumInputLength: finalConfig.searchable ? finalConfig.minimumInputLength : Infinity,
      maximumSelectionLength: finalConfig.maximumSelectionLength,
      closeOnSelect: finalConfig.closeOnSelect,
      tags: finalConfig.tags
    });
  }

  /**
   * Populate a Select2 dropdown with data
   * @param selector - jQuery selector for the dropdown element
   * @param data - Array of values to populate
   * @param placeholderText - Text for the placeholder when no option is selected
   */
  public static populateSelect2(
    selector: string, 
    data: string[], 
    placeholderText?: string
  ): void {
    const options: Select2Option[] = [];
    
    // Add data options (no empty option for placeholder)
    data.forEach(item => {
      if (item && item.trim() !== '') {
        options.push({ id: item, text: item });
      }
    });

    // Re-initialize with new options and placeholder
    const config: Select2Config = {
      placeholder: placeholderText || 'Select an option...',
      allowClear: true
    };
    
    this.initializeSelect2(selector, options, config);
  }

  /**
   * Update Select2 dropdown options dynamically
   * @param selector - jQuery selector for the dropdown element
   * @param options - New array of Select2Option objects
   */
  public static updateSelect2Options(selector: string, options: Select2Option[]): void {
    // Clear existing options
    $(selector).empty();

    // Add new options
    options.forEach(option => {
      const $option = $(`<option value="${option.id}">${option.text}</option>`);
      if (option.selected) {
        $option.prop('selected', true);
      }
      $(selector).append($option);
    });

    // Trigger change to update Select2
    $(selector).trigger('change');
  }

  /**
   * Get the selected value(s) from a Select2 dropdown
   * @param selector - jQuery selector for the dropdown element
   * @returns Selected value or array of values for multi-select
   */
  public static getSelectedValue(selector: string): string | string[] {
    return $(selector).val() as string | string[];
  }

  /**
   * Set the selected value(s) for a Select2 dropdown
   * @param selector - jQuery selector for the dropdown element
   * @param value - Value or array of values to select
   */
  public static setSelectedValue(selector: string, value: string | string[]): void {
    $(selector).val(value).trigger('change');
  }

  /**
   * Clear the selection in a Select2 dropdown
   * @param selector - jQuery selector for the dropdown element
   */
  public static clearSelection(selector: string): void {
    ($(selector) as any).val(null).trigger('change');
  }

  /**
   * Destroy Select2 instance
   * @param selector - jQuery selector for the dropdown element
   */
  public static destroySelect2(selector: string): void {
    if ($(selector).hasClass('select2-hidden-accessible')) {
      ($(selector) as any).select2('destroy');
    }
  }

  /**
   * Apply Select2 styling to all dropdowns in a container
   * @param containerSelector - jQuery selector for the container element
   * @param config - Select2 configuration options
   */
  public static initializeAllDropdowns(containerSelector: string, config: Select2Config = {}): void {
    $(containerSelector).find('select').each((index, element) => {
      const $select = $(element);
      const options: Select2Option[] = [];
      
      // Extract existing options
      $select.find('option').each((i, opt) => {
        const $opt = $(opt);
        options.push({
          id: $opt.val() as string,
          text: $opt.text(),
          selected: $opt.prop('selected')
        });
      });

      // Initialize Select2
      this.initializeSelect2(`#${element.id}`, options, config);
    });
  }
}
