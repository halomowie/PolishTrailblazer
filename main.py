import requests
import pandas as pd
from datetime import datetime

query = """
fragment TrailheadRank on TrailheadRank {
  __typename
  title
  requiredPointsSum
  requiredBadgesCount
  imageUrl
}

fragment PublicProfile on PublicProfile {
  __typename
  trailheadStats {
    __typename
    earnedPointsSum
    earnedBadgesCount
    completedTrailCount
    rank {
      ...TrailheadRank
    }
    nextRank {
      ...TrailheadRank
    }
  }
}

query GetTrailheadRank($slug: String, $hasSlug: Boolean!) {
  profile(slug: $slug) @include(if: $hasSlug) {
    ... on PublicProfile {
      ...PublicProfile
    }
    ... on PrivateProfile {
      __typename
    }
  }
}
"""

url = 'https://profile.api.trailhead.com/graphql'
headers = {'Content-Type': 'application/json'}


def send_request_trailhead(slug):
    variables = {
        "hasSlug": True,
        "slug": slug
    }
    try:
        response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)
        stats_data = response.json()['data']['profile']['trailheadStats']
    except:
        return None

    return stats_data


if __name__ == '__main__':
    df = pd.read_excel('Trailblazers.xlsx', engine='openpyxl')
    df['slugs'] = df['Profile link'].str.split('/').str[-1]


    df['Points'] = 0
    df['Badges'] = 0
    df['Trails'] = 0
    df['Scrap status'] = 'Success'

    result_df = pd.DataFrame()
    for index, row in df.iterrows():
        print(f"Processing {row['slugs']}")
        result = send_request_trailhead(row['slugs'])
        if result == None:
            df.at[index, 'Scrap status'] = 'Error'
        else:
            df.at[index, 'Points'] = result['earnedPointsSum']
            df.at[index, 'Badges'] = result['earnedBadgesCount']
            df.at[index, 'Trails'] = result['completedTrailCount']
            df.at[index, 'Scrap status'] = 'Success'

    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"trailblazers_stats_{current_time}.xlsx"
    df.to_excel(filename, engine='openpyxl', index=False)

