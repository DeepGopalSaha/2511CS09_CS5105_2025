
#include <iostream>
#include <vector>
#include <queue>
#include <climits>
using namespace std;

class Solution {
public:
    int shortestPathBinaryMatrix(vector<vector<int>>& grid) {
        int n = grid.size(), m = grid[0].size();
        vector<vector<int>> dist(n, vector<int>(m, INT_MAX));
        queue<pair<pair<int,int>, int>> q;

        if(n == 1 && grid[0][0] == 0) return 1;
        if(grid[0][0] == 1 || grid[n-1][m-1] == 1) return -1;

        int delrow[] = {-1,-1,-1,0,1,1,1,0};
        int delcol[] = {-1,0,1,1,1,0,-1,-1};
        dist[0][0] = 1;
        q.push({{0,0},1});

        while(!q.empty()){
            auto it = q.front();
            int row = it.first.first;
            int col = it.first.second;
            int d = it.second;
            q.pop();

            for(int i=0;i<8;i++){
                int newr = row + delrow[i];
                int newc = col + delcol[i];

                if(newr>=0 && newr<n && newc>=0 && newc<m && grid[newr][newc]==0 && d+1 < dist[newr][newc]){
                    dist[newr][newc] = d+1;
                    if(newr == n-1 && newc == m-1){
                        return d+1;
                    }
                    q.push({{newr,newc},d+1});
                }
            }
        }
        return -1;
    }
};

int main() {
    Solution sol;

    // Example test case
    vector<vector<int>> grid = {
        {0, 1},
        {1, 0}
    };

    int result = sol.shortestPathBinaryMatrix(grid);
    cout << "Shortest path length: " << result << endl;

    return 0;
}
