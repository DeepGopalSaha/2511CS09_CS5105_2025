#include <iostream>
#include <vector>
#include <queue>
using namespace std;

vector<vector<int>> nearest(vector<vector<int>>& grid) {
    int m = grid.size(), n = grid[0].size();
    
    vector<vector<int>> vis(m, vector<int>(n, 0));
    vector<vector<int>> dist(m, vector<int>(n, 0));
    
    queue<pair<pair<int,int>,int>> q;
    for(int i = 0; i < m; i++) {
        for(int j = 0; j < n; j++) {
            if(grid[i][j] == 1) {
                q.push({{i, j}, 0});
                vis[i][j] = 1;
            }
        }
    }
    
    // BFS
    while(!q.empty()) {
        int r = q.front().first.first;
        int c = q.front().first.second;
        int d = q.front().second;
        q.pop();
        
        int deltarow[] = {-1, 0, 1, 0};
        int deltacol[] = {0, 1, 0, -1};
        
        for(int i = 0; i < 4; i++) {
            int nextrow = r + deltarow[i];
            int nextcol = c + deltacol[i];
            
            if(nextrow >= 0 && nextrow < m && nextcol >= 0 && nextcol < n 
               && vis[nextrow][nextcol] != 1) {
                vis[nextrow][nextcol] = 1;
                dist[nextrow][nextcol] = d + 1;
                q.push({{nextrow, nextcol}, d + 1});
            }
        }
    }
    return dist;
}

int main() {
    vector<vector<int>> grid = {
        {0, 1, 0},
        {0, 0, 0},
        {1, 0, 0}
    };

    vector<vector<int>> ans = nearest(grid);

    cout << "Distance matrix:\n";
    for(auto &row : ans) {
        for(auto &val : row) {
            cout << val << " ";
        }
        cout << "\n";
    }

    return 0;
}
