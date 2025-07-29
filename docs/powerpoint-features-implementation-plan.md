# PowerPoint Dynamic Slide Duplication and Filtering Implementation Plan

## Overview
This document outlines the implementation plan for two new PowerPoint merge features:
1. **Dynamic Slide Duplication**: Create multiple slides from a template slide based on list data
2. **Slide Filtering**: Include/exclude specific slides in the final output

## Feature 1: Dynamic Slide Duplication

### Description
Loop through a list in merge data and create a new slide for each item, merging row fields into the template slide's merge fields. If the list is empty or absent, no slides are added.

### Design Approach
- Add a special merge field syntax to mark "template slides" for duplication
- Use a reserved prefix like `{{#list:listname}}` to identify template slides  
- For each item in the list, duplicate the template slide and merge item data
- Remove the original template slide after duplication

### Implementation Details

#### 1. Configuration Structure
```json
{
  "global_settings": {
    "powerpoint": {
      "dynamic_slides": {
        "enabled": true,
        "template_marker": "{{#list:",
        "remove_template_slides": true
      }
    }
  }
}
```

#### 2. Data Structure Example
```json
{
  "company_name": "Acme Corp",
  "products": [
    {"name": "Product A", "price": "$100", "description": "..."},
    {"name": "Product B", "price": "$200", "description": "..."}
  ]
}
```

#### 3. Template Slide Markup and Merge Field Naming Convention

##### Basic Template Slide
The template slide would contain:
- `{{#list:products}}` - marks this as a template slide for the "products" list
- `{{name}}`, `{{price}}`, `{{description}}` - merge fields for list items (simple names, no prefix needed)

##### Merge Field Naming Convention
When processing template slides, the merger uses **simple field names** within the list context:

**Why Simple Names?**
- Context is already established by `{{#list:products}}`
- Cleaner and more intuitive for users
- Consistent with popular template engines (Handlebars, Mustache, Jinja)
- No need to repeat the list name in every field

**Example:**
```
Instead of: {{products.name}} or {{products[0].name}}
Use: {{name}}
```

##### Special Variables
The following special variables are available within list context:
- `{{$index}}` - Current item index (0-based)
- `{{$position}}` - Current item position (1-based, for display)
- `{{$first}}` - Boolean, true for first item
- `{{$last}}` - Boolean, true for last item
- `{{$odd}}` / `{{$even}}` - For alternating row styles

##### Accessing Parent Context
To access data outside the current list item:
- `{{../company_name}}` - Access parent context (one level up)
- `{{$root.company_name}}` - Access root context from any depth

**Example Template Slide:**
```
{{#list:products}}
Product {{$position}} of {{../total_products}}
Name: {{name}}
Price: {{price}}
Company: {{$root.company_name}}
{{#if $last}}Thank you for your business!{{/if}}
```

##### Nested Object Support
For list items with nested objects:
- `{{contact.email}}` - Access nested field within the current item
- `{{address.city}}` - Still relative to current list item context

## Feature 2: Slide Filtering (Include/Exclude)

### Description
Allow configuration to include only specific slides or exclude specific slides from the final output.

### Design Approach
- Add include/exclude lists to configuration
- Filter slides before processing based on slide numbers
- If include list is present and not empty, include only listed slides
- If exclude list is present and not empty, exclude listed slides
- If both are empty, process all slides normally

### Implementation Details

#### 1. Configuration Structure
```json
{
  "global_settings": {
    "powerpoint": {
      "slide_filter": {
        "include_slides": [1, 3, 5],  // Optional: only include these slides
        "exclude_slides": [2, 4]      // Optional: exclude these slides
      }
    }
  }
}
```

#### 2. Processing Logic
- Include list takes precedence over exclude list
- Slide numbers are 1-based (matching PowerPoint UI)
- Invalid slide numbers are ignored with warnings

## Implementation Steps

### Phase 1: Core Infrastructure (1-2 hours)
1. **Update Configuration Schema**
   - Add `powerpoint` section to `global_settings` in default_config.json
   - Add configuration validation in ConfigManager

2. **Create Slide Utilities Module** (`src/utils/slide_utils.py`)
   - `duplicate_slide()` - Deep copy a slide with all shapes
   - `get_slide_index()` - Convert slide number to index
   - `is_template_slide()` - Check if slide has template marker

### Phase 2: Slide Filtering (1 hour)
1. **Update PowerPointProcessor.merge_data()**
   - Add `_filter_slides()` method before processing
   - Implement include/exclude logic
   - Add logging for filtered slides

2. **Testing**
   - Test with various include/exclude combinations
   - Test with invalid slide numbers
   - Test with empty configuration

### Phase 3: Dynamic Slide Duplication (2-3 hours)
1. **Update PowerPointProcessor**
   - Add `_process_dynamic_slides()` method
   - Implement template slide detection
   - Implement slide duplication logic
   - Update merge field processing for list context

2. **Implement Slide Duplication**
   - Use python-pptx slide duplication approach
   - Preserve formatting and layout
   - Handle images and complex shapes

3. **Merge Field Context Processing**
   - Modify `_process_text_shape()` to accept context parameter
   - Implement `_create_list_context()` for special variables
   - Add `_resolve_field_path()` for parent/root access
   - Update field replacement to use context-aware resolution

4. **Testing**
   - Test with empty lists
   - Test with multiple template slides
   - Test with special variables ($index, $position, etc.)
   - Test parent context access (../field_name)
   - Test root context access ($root.field_name)

### Phase 4: Integration and Edge Cases (1 hour)
1. **Integration**
   - Ensure both features work together
   - Handle template slides in filtered slides
   - Update documentation

2. **Error Handling**
   - Handle missing list data gracefully
   - Validate template slide syntax
   - Add comprehensive logging

## Code Structure Changes

### 1. New Files
- `src/utils/slide_utils.py` - Slide manipulation utilities

### 2. Modified Files
- `src/pptx_processor.py` - Main implementation
- `config/default_config.json` - Add PowerPoint configuration
- `src/config_manager.py` - Add configuration validation

### 3. Key Methods to Add

#### In `pptx_processor.py`:
```python
def _filter_slides(self, slides, config):
    """Filter slides based on include/exclude configuration."""
    
def _process_dynamic_slides(self, data, config):
    """Process template slides and create duplicates from list data."""
    
def _is_template_slide(self, slide, config):
    """Check if slide contains template markers."""
    
def _extract_list_name(self, slide, config):
    """Extract the list name from template marker."""
    
def _duplicate_slide(self, slide, index):
    """Create a duplicate of the slide at the specified index."""
    
def _create_list_context(self, item, index, total, parent_data):
    """Create context for list item with special variables.
    
    Returns dict with:
    - All item fields
    - Special variables: $index, $position, $first, $last, $odd, $even
    - Parent context access via ../field_name
    - Root context access via $root
    """
    
def _resolve_field_path(self, field_path, context):
    """Resolve field path including parent context (../) and root ($root)."""
```

## Testing Plan

### Unit Tests
1. Test slide filtering logic
2. Test template slide detection
3. Test merge field replacement in duplicated slides

### Integration Tests
1. Test complete merge with dynamic slides
2. Test with various data structures
3. Test error scenarios

### Manual Testing Checklist
- [ ] Create sample template with template slides
- [ ] Test with empty list data
- [ ] Test with multiple items in list
- [ ] Test include/exclude filters
- [ ] Test combined features
- [ ] Test with complex slides (images, tables)

## Rollback Plan
Both features are controlled by configuration flags:
- `global_settings.powerpoint.dynamic_slides.enabled`
- `global_settings.powerpoint.slide_filter` (empty = disabled)

Setting these to false/empty will disable the features without code changes.

## Timeline Estimate
- **Total Implementation**: 5-7 hours
- **Testing**: 2-3 hours
- **Documentation**: 1 hour

## Risks and Mitigations
1. **Risk**: python-pptx slide duplication complexity
   - **Mitigation**: Start with simple implementation, add complexity incrementally

2. **Risk**: Performance with large lists
   - **Mitigation**: Add configurable limits, optimize duplication

3. **Risk**: Breaking existing functionality
   - **Mitigation**: Feature flags, comprehensive testing

## Success Criteria
1. Template slides can be duplicated based on list data
2. Slides can be filtered using include/exclude lists
3. No regression in existing merge functionality
4. Clear error messages for configuration issues
5. Performance remains acceptable (<5s for typical presentations)

## Future Enhancements
1. Support for nested lists
2. Conditional slide inclusion based on data
3. Slide reordering capability
4. Template slide variants based on data type

## Complete Example

### Input Data:
```json
{
  "company_name": "TechCorp Solutions",
  "total_products": 3,
  "year": 2024,
  "products": [
    {
      "name": "CloudSync Pro",
      "price": "$299",
      "description": "Enterprise cloud synchronization",
      "features": {
        "storage": "Unlimited",
        "users": "500+"
      }
    },
    {
      "name": "DataVault",
      "price": "$199",
      "description": "Secure data backup solution",
      "features": {
        "storage": "10TB",
        "users": "100"
      }
    },
    {
      "name": "TeamHub",
      "price": "$99",
      "description": "Team collaboration platform",
      "features": {
        "storage": "1TB",
        "users": "50"
      }
    }
  ]
}
```

### Template Slide Content:
```
{{#list:products}}
┌─────────────────────────────────────────┐
│ {{$root.company_name}} Product Catalog  │
│ Item {{$position}} of {{../total_products}} │
├─────────────────────────────────────────┤
│ Product: {{name}}                       │
│ Price: {{price}}                        │
│                                         │
│ {{description}}                         │
│                                         │
│ Features:                               │
│ • Storage: {{features.storage}}         │
│ • Users: {{features.users}}             │
│                                         │
│ © {{$root.year}} {{$root.company_name}} │
└─────────────────────────────────────────┘
```

### Result:
Creates 3 slides with the template content filled with each product's data, with proper context resolution for nested fields, parent context, and special variables.