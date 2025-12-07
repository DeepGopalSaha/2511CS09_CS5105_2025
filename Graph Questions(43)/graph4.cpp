#include <bits/stdc++.h>
using namespace std;
vector<vector<int>> nearestOnes(vector<vector<int>>& mat) {
    int n=mat.size(), m=mat[0].size();
    vector<vector<int>> dist(n, vector<int>(m, -1));
    queue<pair<int,int>>q;
    for(int i=0;i<n;i++) for(int j=0;j<m;j++) if(mat[i][j]==1){ dist[i][j]=0; q.push({i,j}); }
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty()){
        auto [r,c]=q.front(); q.pop();
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            if(dist[nr][nc]==-1){ dist[nr][nc]=dist[r][c]+1; q.push({nr,nc}); }
        }
    }
    return dist;
}
int main(){ return 0; }
