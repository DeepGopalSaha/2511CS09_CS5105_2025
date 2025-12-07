#include <bits/stdc++.h>
using namespace std;
void brute_next_permutation(vector<int>& a){ next_permutation(a.begin(), a.end()); }
void optimal_next_permutation(vector<int>& a){
    int n=a.size(), i=n-2;
    while(i>=0 && a[i]>=a[i+1]) i--;
    if(i>=0){
        int j=n-1;
        while(a[j]<=a[i]) j--;
        swap(a[i], a[j]);
    }
    reverse(a.begin()+i+1, a.end());
}
int main(){return 0;}