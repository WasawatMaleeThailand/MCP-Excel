# client.py (เฉพาะส่วนที่แก้ไข)
# ... (โค้ดส่วนอื่นเหมือนเดิม)

async def main():
    # ... (โค้ดส่วนอื่นเหมือนเดิม)

    # --- สร้าง Payload สำหรับเรียกใช้ write_data_to_excel ---
    payload = [{
        "type": "callTool",
        "id": f"call-{uuid.uuid4()}",
        "name": "write_data_to_excel",
        "arguments": {
            "filepath": "sales_report.xlsx",
            "sheet_name": "Quarter1",
            "data": [
                ["Product", "Units Sold", "Price"],
                ["Laptop", 20, 35000],
                ["Mouse", 150, 750],
                ["Keyboard", 100, 1200]
            ],
            "start_cell": "A1"
        }
    }]

    # --- ยิง request ไปที่ Server ---
    print("\nCalling 'write_data_to_excel' tool...")
    response_frames = await session.call_tool(
        name=payload[0]['name'],
        arguments=payload[0]['arguments']
    )
    print("Response from server:")
    print(response_frames)

# ... (โค้ดส่วน if __name__ == "__main__": เหมือนเดิม)
