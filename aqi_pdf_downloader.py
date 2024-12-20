import os
import requests
from datetime import datetime, timedelta
import logging
from pathlib import Path
from tqdm import tqdm

class AQIBulletinDownloader:
    """Downloads AQI bulletins from CPCB website for a specified date range."""
    
    def __init__(self, output_dir="aqi_bulletins"):
        """
        Initialize the downloader with output directory
        
        Args:
            output_dir (str): Directory where PDFs will be stored
        """
        self.base_url = "https://cpcb.nic.in/upload/Downloads"
        self.output_dir = Path(output_dir)
        self.setup_logging()
        
    def setup_logging(self):
        """Configure logging for the downloader"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('aqi_download.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def create_output_directory(self):
        """Create output directory if it doesn't exist"""
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def get_date_range(self, start_date_str, end_date_str):
        """
        Generate a list of dates between start and end dates
        
        Args:
            start_date_str (str): Start date in YYYY-MM-DD format
            end_date_str (str): End date in YYYY-MM-DD format
            
        Returns:
            list: List of datetime objects
        """
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            
            if end_date < start_date:
                raise ValueError("End date must be after start date")
                
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                date_list.append(current_date)
                current_date += timedelta(days=1)
            return date_list
            
        except ValueError as e:
            self.logger.error(f"Date format error: {e}")
            raise
    
    def download_bulletin(self, date):
        """
        Download AQI bulletin for a specific date
        
        Args:
            date (datetime): Date to download bulletin for
            
        Returns:
            bool: True if download successful, False otherwise
        """
        date_str = date.strftime("%Y%m%d")
        filename = f"AQI_Bulletin_{date_str}.pdf"
        url = f"{self.base_url}/{filename}"
        output_path = self.output_dir / filename
        
        if output_path.exists():
            self.logger.info(f"File already exists: {filename}")
            return True
        
        try:
            # Try first URL pattern
            response = requests.get(url)
            
            # If first pattern fails, try alternate URL pattern
            if response.status_code != 200:
                alternate_url = f"{self.base_url}/AQI_Bulletin_{date_str}.pdf"
                response = requests.get(alternate_url)
                if response.status_code != 200:
                    self.logger.warning(f"Bulletin not found for date: {date_str}")
                    return False
            
            # Verify PDF content type
            content_type = response.headers.get('content-type', '').lower()
            if 'application/pdf' not in content_type and 'binary/octet-stream' not in content_type:
                self.logger.warning(f"Not a PDF file for date: {date_str}")
                return False
            
            total_size = int(response.headers.get('content-length', 0))
            
            with open(output_path, 'wb') as file, tqdm(
                desc=filename,
                total=total_size,
                unit='iB',
                unit_scale=True,
                unit_divisor=1024,
            ) as pbar:
                for data in response.iter_content(chunk_size=1024):
                    size = file.write(data)
                    pbar.update(size)
            
            self.logger.info(f"Successfully downloaded: {filename}")
            return True
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Failed to download {filename}: {e}")
            if output_path.exists():
                output_path.unlink()
            return False
    
    def download_range(self, start_date, end_date):
        """
        Download bulletins for a date range
        
        Args:
            start_date (str): Start date in YYYY-MM-DD format
            end_date (str): End date in YYYY-MM-DD format
            
        Returns:
            tuple: (successful_downloads, failed_downloads)
        """
        self.create_output_directory()
        dates = self.get_date_range(start_date, end_date)
        
        successful = 0
        failed = 0
        
        self.logger.info(f"Starting download for date range: {start_date} to {end_date}")
        
        for date in tqdm(dates, desc="Overall Progress"):
            if self.download_bulletin(date):
                successful += 1
            else:
                failed += 1
        
        return successful, failed

def main():
    """Main function to run the downloader"""
    try:
        print("CPCB AQI Bulletin Downloader")
        print("----------------------------")
        print("Downloads AQI bulletins from https://cpcb.nic.in/upload/Downloads/")
        print("Format: AQI_Bulletin_YYYYMMDD.pdf")
        print("----------------------------\n")
        
        # Get user input
        start_date = input("Enter start date (YYYY-MM-DD): ")
        end_date = input("Enter end date (YYYY-MM-DD): ")
        output_dir = input("Enter output directory (default: aqi_bulletins): ").strip() or "aqi_bulletins"
        
        # Initialize and run downloader
        downloader = AQIBulletinDownloader(output_dir)
        successful, failed = downloader.download_range(start_date, end_date)
        
        print(f"\nDownload Summary:")
        print(f"Successful downloads: {successful}")
        print(f"Failed downloads: {failed}")
        print(f"Files saved in: {output_dir}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        logging.error(f"Program error: {e}", exc_info=True)

if __name__ == "__main__":
    main()
