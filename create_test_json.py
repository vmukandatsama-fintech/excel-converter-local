"""
Create a test JSON file to trigger PDF generation
"""
import json
import os
from pathlib import Path
from datetime import datetime

def create_test_json():
    # Define the path
    user_profile = Path(os.environ['USERPROFILE'])
    json_dir = user_profile / "Premium Leaf Zimbabwe" / "v2SavannahTEST - Collection Vouchers" / "Json"
    
    # Ensure directory exists
    json_dir.mkdir(parents=True, exist_ok=True)
    
    # Create test data
    test_data = {
        "FileName": f"TEST-CV-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
        "Header": {
            "Requestor": "Victor Mukandatsama",
            "Date": datetime.now().strftime("%Y-%m-%d"),
            "Approval": "John Manager",
            "Authorization": "Jane Director",
            "Comments": "Test Collection Voucher - Email Functionality Test",
            "Driver": "Test Driver",
            "DriverID": "DRV-001",
            "Truck": "TRK-123",
            "Trailer": "TRL-456",
            "Farmer": "Test Farmer Ltd"
        },
        "Lines": [
            {
                "Line": 1,
                "UOM": "KG",
                "Item": "Premium Tobacco Leaf",
                "Requested": 1000,
                "Issue": 1000,
                "AlreadyIssued": 0,
                "Balance": 0
            },
            {
                "Line": 2,
                "UOM": "KG",
                "Item": "Standard Tobacco Leaf",
                "Requested": 500,
                "Issue": 500,
                "AlreadyIssued": 0,
                "Balance": 0
            },
            {
                "Line": 3,
                "UOM": "EA",
                "Item": "Packaging Materials",
                "Requested": 50,
                "Issue": 50,
                "AlreadyIssued": 0,
                "Balance": 0
            }
        ]
    }
    
    # Generate filename
    filename = f"{test_data['FileName']}.json"
    json_path = json_dir / filename
    
    # Write JSON file
    with open(json_path, 'w') as f:
        json.dump(test_data, f, indent=2)
    
    print("=" * 60)
    print("TEST JSON FILE CREATED")
    print("=" * 60)
    print(f"File: {json_path}")
    print(f"Voucher: {test_data['FileName']}")
    print(f"\nContent:")
    print(json.dumps(test_data, indent=2))
    print("\n" + "=" * 60)
    print("The monitoring application should detect this file")
    print("and automatically generate a PDF + send email.")
    print("=" * 60)
    
    return json_path

if __name__ == "__main__":
    try:
        json_path = create_test_json()
        print(f"\n✓ Test JSON created successfully!")
        print(f"\nIf the main application is running, it will:")
        print(f"  1. Detect the file")
        print(f"  2. Generate PDF from template")
        print(f"  3. Send email to configured recipients")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
