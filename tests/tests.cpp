#include "officeread/officeread.hpp"

#include <cassert>
#include <iostream>

int main() {
  officeread::Result r;
  r.text = "hello";
  r.images.push_back({"a b.png", "picture", ".png", {}});
  const auto md = r.markdown("images");
  assert(md.find("hello") != std::string::npos);
  assert(md.find("a%20b.png") != std::string::npos);
  std::cout << "ok\n";
}
