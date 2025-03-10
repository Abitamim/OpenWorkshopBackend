from fastapi import FastAPI, Form, Request, status, File, UploadFile
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import uvicorn
from datetime import datetime
from io import BytesIO


app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    print('Request for index page received')
    return templates.TemplateResponse('index.html', {"request": request})
    
@app.get('/admin', response_class=HTMLResponse)
async def admin(request: Request):
    print('Request for admin page recieved')
    return templates.TemplateResponse('admin.html', {"request": request})

@app.post('/admin', response_class=HTMLResponse)
async def handle_upload(request: Request, file: UploadFile = File(...)):
    try:
        # Read the Excel file
        contents = await file.read()
        
        # You can process the Excel file here
        df = pd.read_excel(BytesIO(contents))
        print("after read")
        df.columns = df.columns.str.strip()
        print('1')
        df['Start Date'] = pd.to_datetime(df['StartDate'], format='%Y-%m-%d')
        df['End Date'] = pd.to_datetime(df['EndDate'], format='%Y-%m-%d')
        print('2')
        df['Workshop Title'] = df['WorkshopTitle']
        print('3')
        df['Duration (Days)'] = (df['End Date'] - df['Start Date']).dt.days + 1
        print(df)
        df['Dates Available'] = df['Start Date'].dt.strftime('%Y-%m-%d')
        grouped = df.groupby('Workshop Title').agg({
            'Duration (Days)': 'first',  # We can just take the first occurrence, assuming length is the same for each group
            'Dates Available': lambda x: ', '.join(sorted(x))  # Combine start dates into a comma-separated list
        }).reset_index()
        
        print("before sort_dates")

        def sort_dates(dates_str):
            # Split the string by commas
            dates_list = dates_str.split(',')
            
            # Convert each date string to a datetime object
            dates_list = [datetime.strptime(date.strip(), '%Y-%m-%d') for date in dates_list]
            
            # Sort the dates in ascending order (closest first)
            dates_list.sort()
        
            # Convert back to string format dd-mmm-yy and join by commas
            return ', '.join([datetime.strftime(date, '%b-%d-%Y') for date in dates_list])
        
        grouped['Dates Available'] = grouped['Dates Available'].apply(sort_dates)
        print("before save")
        with pd.ExcelWriter("workshops_data/clean_workshops.xlsx") as writer:
            grouped.to_excel(writer, index=False, sheet_name='Workshops')
        
        return templates.TemplateResponse('admin.html', {
            "request": request,
            "message": f"Successfully processed {file.filename}",
            "success": True
        })
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return templates.TemplateResponse('admin.html', {
            "request": request,
            "message": f"Error processing file: {str(e)}",
            "success": False
        })

if __name__ == '__main__':
    uvicorn.run('myapp:app', host='0.0.0.0', port=8000)

