import numpy as np
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
    response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)
    stats_data = response.json()['data']['profile']['trailheadStats']
    print(stats_data)

    return stats_data


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    df = pd.read_excel('Trailblazers.xlsx', engine='openpyxl')
    df['slugs'] = df['Profile link'].str.split('/').str[-1]

    df['Points'] = 0
    df['Badges'] = 0
    df['Trails'] = 0
    # df['']

    # print(df)

    result_df = pd.DataFrame()
    for index, row in df.iterrows():
        print(row['slugs'])
        result = send_request_trailhead(row['slugs'])
        df.at[index, 'Points'] = result['earnedPointsSum']
        df.at[index, 'Badges'] = result['earnedBadgesCount']
        df.at[index, 'Trails'] = result['completedTrailCount']

    print(df)
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"trailblazers_stats_{current_time}.xlsx"
    df.to_excel(filename, engine='openpyxl', index=False)

