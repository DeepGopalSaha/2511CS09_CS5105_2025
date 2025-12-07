#include <bits/stdc++.h>
using namespace std;
void brute_rotate_matrix(vector<vector<int>>& a){
    int n=a.size();
    vector<vector<int>> b=a;
    for(int i=0;i<n;++i) for(int j=0;j<n;++j) a[j][n-1-i]=b[i][j];
}
void optimal_rotate_matrix(vector<vector<int>>& a){
    int n=a.size();
    for(int i=0;i<n;++i) for(int j=i+1;j<n;++j) swap(a[i][j], a[j][i]);
    for(int i=0;i<n;++i) reverse(a[i].begin(), a[i].end());
}
int main(){return 0;}