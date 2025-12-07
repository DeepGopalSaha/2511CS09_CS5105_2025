#include <bits/stdc++.h>
using namespace std;
vector<vector<int>> floodFill(vector<vector<int>>& image, int sr, int sc, int newColor){
    int n=image.size(), m=image[0].size();
    int orig = image[sr][sc];
    if(orig==newColor) return image;
    queue<pair<int,int>>q; q.push({sr,sc}); image[sr][sc]=newColor;
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty()){
        auto [r,c]=q.front(); q.pop();
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            if(image[nr][nc]==orig){ image[nr][nc]=newColor; q.push({nr,nc}); }
        }
    }
    return image;
}
int main(){ return 0; }
