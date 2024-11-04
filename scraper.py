import pandas as pd
from requests_html import HTMLSession
from bs4 import BeautifulSoup as bs
from pytube import YouTube 
import os


# Initialize HTML session
session = HTMLSession()

def get_video_metadata(url):
    # Fetch and parse the HTML content of the YouTube video page
    response = session.get(url)
    response.html.render(sleep=1)
    soup = bs(response.html.html, "html.parser")
    yt = YouTube(url)

    download_path='.\\videos'
    video = yt.streams.filter(file_extension='mp4', res='720p').first()

    # Extract video metadata
    video_meta = {}
    video_meta["title"] = yt.title
    video_meta["view"] = soup.find("meta", itemprop="interactionCount")["content"]
    video_meta["tags"] = ', '.join([meta.attrs.get("content") for meta in soup.find_all("meta", {"property": "og:video:tag"})])
    video_meta["description"] = soup.find("div", id="description").text.strip().replace('\n', ' ')
    
    # Extract published date
    video_meta["date_published"] = yt.publish_date
    replase=soup.find("meta", property="og:title")["content"].replace('|', '')
    video_path = os.path.join(download_path, f"{replase}.mp4")  # Change the path as needed
    video.download(download_path)
    # Extract likes and Size
    video_meta["Length"] = yt.length
    video_meta["Size"] = (os.path.getsize(video_path))*0.000001
    print(video_meta)
    return video_meta

# Read CSV file with YouTube video links and existing metadata
df = pd.read_excel(r'test.xlsx')

# Change the data types of the columns
df['Title'] = df['Title'].astype('object')
df['View'] = df['View'].astype('Int64')  # Use 'Int64' (capital I) to handle NaN values
df['Tags'] = df['Tags'].astype('object')
df['Description'] = df['Description'].astype('object')
df['Date Published'] = df['Date Published'].astype('object')
df['Length'] = df['Length'].astype('Int64')
df['Size'] = df['Size'].astype('float64')

# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    video_link = row['Video Links']
    try:
        # Get metadata for each video link
        video_metadata = get_video_metadata(video_link)
        
        # Update the DataFrame with the extracted metadata
        df.at[index, 'Title'] = video_metadata["title"]
        df.at[index, 'View'] = int(video_metadata["view"].replace(',', '')) if video_metadata["view"] else None
        df.at[index, 'Tags'] = video_metadata["tags"]
        df.at[index, 'Description'] = video_metadata["description"]
        df.at[index, 'Date Published'] = video_metadata["date_published"]
        df.at[index, 'Length'] = video_metadata["Length"]
        df.at[index, 'Size'] = video_metadata["Size"]
        
    except Exception as e:
        print(f"Error processing {video_link}: {str(e)}")

# Save the updated DataFrame to the original CSV file
df.to_excel(r'test.xlsx', index=False)
print("CSV file has been updated.")