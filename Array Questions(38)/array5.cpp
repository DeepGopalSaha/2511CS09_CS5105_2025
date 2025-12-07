#include <bits/stdc++.h>
using namespace std;
void brute_left_rotate_one(vector<int>& a){ if(a.empty()) return; int first=a[0]; for(int i=0;i+1<a.size();++i) a[i]=a[i+1]; a.back()=first; }
void optimal_left_rotate_one(vector<int>& a){ if(a.empty()) return; rotate(a.begin(), a.begin()+1, a.end()); }
int main(){return 0;}