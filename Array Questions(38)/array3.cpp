#include <bits/stdc++.h>
using namespace std;
bool brute_is_sorted(vector<int>& a){ for(int i=1;i<a.size();++i) if(a[i]<a[i-1]) return false; return true;}
bool optimal_is_sorted(vector<int>& a){ return is_sorted(a.begin(), a.end()); }
int main(){return 0;}