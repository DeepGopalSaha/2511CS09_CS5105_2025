#include <bits/stdc++.h>
using namespace std;
int brute_max_consecutive_ones(vector<int>& a){ int best=0; for(int i=0;i<a.size();++i){ int cur=0; for(int j=i;j<a.size() && a[j]==1;++j) cur++; best=max(best,cur);} return best;}
int optimal_max_consecutive_ones(vector<int>& a){ int cur=0,best=0; for(int x:a){ if(x==1) cur++; else cur=0; best=max(best,cur);} return best;}
int main(){return 0;}