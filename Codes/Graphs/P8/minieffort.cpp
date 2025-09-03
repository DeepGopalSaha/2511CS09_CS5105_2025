
#include <iostream>
#include <vector>
#include <queue>
#include <climits>
using namespace std;

class Solution {
public:
    int minimumEffortPath(vector<vector<int>>& heights) {
        priority_queue<pair<int, pair<int,int>>, 
                       vector<pair<int, pair<int,int>>>, 
                       greater<pair<int, pair<int,int>>>> pq;
        
        int n = heights.size(), m = heights[0].size();
        vector<vector<int>> dist(n, vector<int>(m, INT_MAX));
        dist[0][0] = 0;
        pq.push({0,{0,0}});

        int delrow[] = {-1,0,1,0};
        int delcol[] = {0,1,0,-1};

        while(!pq.empty()){
            auto it = pq.top();
            int r = it.second.first, c = it.second.second, d = it.first;
            pq.pop();

            if(r == n-1 && c == m-1) return d;

            for(int i=0;i<4;i++){
                int nr = r + delrow[i];
                int nc = c + delcol[i];

                if(nr>=0 && nr<n && nc>=0 && nc<m){
                    int effort = max(abs(heights[nr][nc]-heights[r][c]), d);
                    if(effort < dist[nr][nc]){
                        dist[nr][nc] = effort;
                        pq.push({effort,{nr,nc}});
                    }
                }
            }
        }
        return 0;
    }
};

int main() {
    Solution sol;

    vector<vector<int>> heights = {
        {1,2,2},
        {3,8,2},
        {5,3,5}
    };

    cout << "Minimum effort path: " << sol.minimumEffortPath(heights) << endl;

    return 0;
}
