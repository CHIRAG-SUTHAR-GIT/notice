# API Integration Documentation

This document explains how the IFSC District Splitter integrates with external APIs to determine bank branch districts.

## Overview

The application uses two APIs with a fallback mechanism:

1. **Razorpay IFSC API** (Primary) - Converts IFSC codes to bank branch details including district
2. **Indian Postal Code API** (Fallback) - Converts pincodes to location details including district

## 1. Razorpay IFSC API

### Purpose
Retrieve detailed information about bank branches using their IFSC (Indian Financial System Code).

### API Details

**Base URL:** `https://ifsc.razorpay.com`

**Endpoint:** `/search`

**Method:** GET

**Query Parameters:**
```
code={IFSC_CODE}    - The IFSC code to search for (required)
limit={number}      - Number of results (optional)
offset={number}     - Pagination offset (optional)
state={state_code}  - Filter by state ISO3166 code (optional)
city={city_name}    - Filter by city name (optional)
bankcode={code}     - Filter by bank code (optional)
```

### Example Requests

**Request 1: Search by IFSC code**
```
GET https://ifsc.razorpay.com/search?code=AUBL0002567
```

**Response:**
```json
{
    "IFSC": "AUBL0002567",
    "BANK": "Axis Bank",
    "BANKCODE": "AUBL",
    "BRANCH": "Ahmedabad",
    "STATE": "GUJARAT",
    "ISO3166": "IN-GJ",
    "DISTRICT": "AHMEDABAD",
    "CENTRE": "AHMEDABAD",
    "CITY": "AHMEDABAD",
    "ADDRESS": "123, Main Street, Ahmedabad",
    "CONTACT": "+91-79-xxxxxxxx",
    "MICR": "380001234",
    "RTGS": true,
    "NEFT": true,
    "IMPS": true,
    "UPI": true,
    "SWIFT": ""
}
```

### Response Fields Used

| Field | Description | Used For |
|-------|-------------|----------|
| DISTRICT | District name | Primary district source |
| STATE | State name | Validation if ISO3166 not present |
| ISO3166 | ISO 3166-2 state code | Validation for Gujarat (IN-GJ) |
| IFSC | IFSC code | Record identifier |

### Data Quality Notes

- IFSC codes must be exactly 11 characters
- Not all bank branches are in the database
- State information is sometimes inconsistent (use ISO3166 when available)
- Some older branches might not be present

### Rate Limits

- Typically 100-1000 requests per minute
- No authentication required
- API is publicly accessible

### Example Code (Python)

```python
import requests

def get_district_from_ifsc(ifsc_code):
    """Fetch district information from IFSC API"""
    try:
        url = f"https://ifsc.razorpay.com/search?code={ifsc_code}"
        response = requests.get(url, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            if data:
                return {
                    'district': data.get('DISTRICT'),
                    'state': data.get('STATE'),
                    'iso3166': data.get('ISO3166')
                }
        return None
    except requests.exceptions.RequestException as e:
        print(f"API Error: {e}")
        return None

# Usage
result = get_district_from_ifsc("AUBL0002567")
print(result)
# Output: {'district': 'AHMEDABAD', 'state': 'GUJARAT', 'iso3166': 'IN-GJ'}
```

## 2. Indian Postal Code API

### Purpose
Retrieve location details (state, district, city) using Indian pincodes.

### API Details

**Base URL:** `https://indian-postal-code-api.herokuapp.com/api/pin`

**Endpoint:** `/find/{PINCODE}`

**Method:** GET

**URL Parameters:**
```
PINCODE - 6-digit Indian postal code (required)
```

### Example Requests

**Request 1: Search by pincode**
```
GET https://indian-postal-code-api.herokuapp.com/api/pin/find/380001
```

**Response:**
```json
[
    {
        "PostOfficeName": "Ahmedabad",
        "Pincode": "380001",
        "District": "Ahmedabad",
        "State": "GUJARAT",
        "Region": "Ahmedabad HO",
        "Longitude": "72.5479",
        "Latitude": "23.1815",
        "Circle": "Gujarat"
    }
]
```

### Response Fields Used

| Field | Description | Used For |
|-------|-------------|----------|
| District | District name | Secondary district source |
| State | State name | Validation for Gujarat |
| Pincode | Postal code | Record identifier |

### Data Structure Notes

- Returns an array of objects (usually 1 entry per pincode)
- All pincodes return at least one entry
- Multiple entries might exist for same pincode (different postal zones)
- Uses first entry in array

### Rate Limits

- Typically 100-500 requests per minute
- No authentication required
- API is publicly accessible

### Example Code (Python)

```python
import requests

def get_district_from_pincode(pincode):
    """Fetch district information from Postal Code API"""
    try:
        url = f"https://indian-postal-code-api.herokuapp.com/api/pin/find/{pincode}"
        response = requests.get(url, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            if data and len(data) > 0:
                return {
                    'district': data[0].get('District'),
                    'state': data[0].get('State')
                }
        return None
    except requests.exceptions.RequestException as e:
        print(f"API Error: {e}")
        return None

# Usage
result = get_district_from_pincode("380001")
print(result)
# Output: {'district': 'Ahmedabad', 'state': 'GUJARAT'}
```

## District Lookup Flow

The application implements a priority-based lookup system:

### Flowchart

```
┌─────────────────────────┐
│   Record from Excel     │
│ (IFSC + Pincode)       │
└──────────┬──────────────┘
           │
           ▼
┌─────────────────────────┐
│ Try IFSC API lookup     │
└──────────┬──────────────┘
           │
        ┌──┴──┐
        │     │
       YES   NO
        │     │
        ▼     │
    ┌─────────┴─────────┐
    │ Got District?     │
    │ Is Gujarat?       │
    │ Valid District?   │
    └────┬────────┬─────┘
         │        │
        YES      NO
         │        │
         ▼        ▼
    ┌────────┐  ┌──────────────────┐
    │ Accept │  │ Try Pincode API  │
    │ Record │  │ Lookup           │
    └────────┘  └────────┬─────────┘
                         │
                      ┌──┴──┐
                      │     │
                     YES   NO
                      │     │
                      ▼     │
                  ┌─────────┴─────────┐
                  │ Got District?     │
                  │ Is Gujarat?       │
                  │ Valid District?   │
                  └────┬────────┬─────┘
                       │        │
                      YES      NO
                       │        │
                       ▼        ▼
                  ┌────────┐  ┌─────────┐
                  │ Accept │  │  Error  │
                  │ Record │  │ Logging │
                  └────────┘  └─────────┘
```

### Validation Steps

1. **IFSC Lookup**
   - Query IFSC API with provided IFSC code
   - Extract DISTRICT, STATE, ISO3166
   - Validate: ISO3166 == "IN-GJ" OR STATE contains "GUJARAT"
   - Validate: DISTRICT is in VALID_DISTRICTS list
   - If valid, accept and store district

2. **Pincode Fallback**
   - Query Postal API with provided pincode
   - Extract District, State
   - Validate: STATE contains "GUJARAT"
   - Validate: District is in VALID_DISTRICTS list
   - If valid, accept and store district

3. **Error Handling**
   - If both lookups fail or validate fails, record as error
   - Error logged with IFSC, pincode, row number
   - Record skipped from output

## Error Handling

### API Errors

```python
try:
    response = requests.get(url, timeout=10)
    
    if response.status_code != 200:
        # API error
        return None
        
    data = response.json()
    if not data:
        # No data returned
        return None
        
except requests.exceptions.Timeout:
    # API took too long
    return None
except requests.exceptions.ConnectionError:
    # Network error
    return None
except Exception as e:
    # Other errors
    return None
```

### Validation Errors

```python
# Check if district matches Gujarat
if state_code != "IN-GJ" and "GUJARAT" not in state:
    return None  # Not from Gujarat

# Check if district is valid
if district not in VALID_DISTRICTS:
    return None  # Unknown district

# Check if data is present
if not district or not state:
    return None  # Missing required fields
```

## Performance Considerations

### Optimization

1. **Batch Processing**
   - Process multiple records sequentially
   - API calls are I/O bound, not CPU bound

2. **Timeouts**
   - 10 second timeout per API call
   - Prevents hanging on slow connections

3. **Caching** (Optional)
   - Can implement local cache for frequently used codes
   - Reduces API calls for duplicate entries

4. **Parallel Processing** (Future)
   - Could use threading for faster processing
   - Must respect API rate limits

### Typical Processing Times

- Per IFSC lookup: 200-500ms
- Per Pincode lookup: 200-500ms
- Per record: 400-1000ms (both lookups)
- 100 records: ~40-100 seconds

## Debugging

### Enable Logging

```python
import logging

logging.basicConfig(level=logging.DEBUG)

# Will show:
# - API URLs and parameters
# - Response status codes
# - Extracted values
# - Validation decisions
```

### Test Queries

```bash
# Test IFSC API
curl "https://ifsc.razorpay.com/search?code=AUBL0002567"

# Test Postal API
curl "https://indian-postal-code-api.herokuapp.com/api/pin/find/380001"
```

## References

- **Razorpay IFSC API**: https://github.com/razorpay/ifsc-api
- **Razorpay IFSC Data**: https://github.com/razorpay/ifsc (CSV download available)
- **Indian Postal Code API**: https://github.com/nstack-in/indian-postal-code-api
- **ISO 3166-2**: https://en.wikipedia.org/wiki/ISO_3166-2:IN

## FAQ

**Q: What if IFSC code is invalid?**
A: The app will try pincode lookup. If that also fails, record is marked as error.

**Q: Can I use only IFSC without pincode?**
A: Yes, but fallback won't work. Records without district will be marked as errors.

**Q: Can I use only pincode without IFSC?**
A: Yes, pincode alone is sufficient for district lookup.

**Q: What if API is down?**
A: The app will show timeout error. Files are not processed. Try again later.

**Q: How accurate are the APIs?**
A: Generally accurate, but occasional inconsistencies exist, especially in district naming.

**Q: Can I modify the valid districts list?**
A: Yes, edit VALID_DISTRICTS in app.py or config.py

**Q: What if branch is in multiple districts?**
A: API returns single primary district. Pincode lookup used as fallback.
