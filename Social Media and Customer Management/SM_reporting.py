import requests

def get_instagram_profile_stats(access_token, instagram_business_account_id):
    api_version = 'v14.0'
    
    # Retrieve general profile information
    profile_url = f'https://graph.facebook.com/{api_version}/{instagram_business_account_id}?fields=business_discovery.username(your_business_username){{username,followers_count,media_count,profile_picture_url,website,biography}}&access_token={access_token}'
    response = requests.get(profile_url)
    profile_data = response.json()
    
    # Retrieve recent media posts
    media_url = f'https://graph.facebook.com/{api_version}/{instagram_business_account_id}/media?fields=media_type,comments_count,like_count,caption,permalink&access_token={access_token}'
    response = requests.get(media_url)
    media_data = response.json()
    
    if 'error' in profile_data:
        error_message = profile_data['error']['message']
        raise Exception(f'Error retrieving Instagram profile statistics: {error_message}')
    
    profile_stats = profile_data['business_discovery']['your_business_username']
    profile_stats['recent_media'] = media_data['data']
    
    return profile_stats

# Replace with your access token and Instagram business account ID
access_token = 'YOUR_ACCESS_TOKEN'
instagram_business_account_id = 'YOUR_INSTAGRAM_BUSINESS_ACCOUNT_ID'

try:
    profile_stats = get_instagram_profile_stats(access_token, instagram_business_account_id)
    
    # Extract required information from the response
    username = profile_stats['username']
    followers_count = profile_stats['followers_count']
    media_count = profile_stats['media_count']
    profile_picture_url = profile_stats['profile_picture_url']
    website = profile_stats['website']
    biography = profile_stats['biography']
    recent_media = profile_stats['recent_media']
    
    # Display the profile statistics
    print(f"Username: {username}")
    print(f"Followers: {followers_count}")
    print(f"Media Count: {media_count}")
    print(f"Profile Picture URL: {profile_picture_url}")
    print(f"Website: {website}")
    print(f"Biography: {biography}")
    
    # Display recent media posts
    print("\nRecent Media:")
    for post in recent_media:
        post_type = post['media_type']
        comments_count = post['comments_count']
        like_count = post['like_count']
        caption = post['caption']
        permalink = post['permalink']
        
        print(f"Type: {post_type}")
        print(f"Comments: {comments_count}")
        print(f"Likes: {like_count}")
        print(f"Caption: {caption}")
        print(f"Permalink: {permalink}")
        print("--------------------------")
        
except Exception as e:
    print(f"An error occurred: {str(e)}")