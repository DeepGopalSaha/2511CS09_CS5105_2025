#include <bits/stdc++.h>
using namespace std;
int brute_max_element(vector<int>& a){int mx=INT_MIN; for(int x:a) mx=max(mx,x); return mx;}
int optimal_max_element(vector<int>& a){return *max_element(a.begin(), a.end());}
int main(){return 0;}